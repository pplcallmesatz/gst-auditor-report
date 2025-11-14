<?php
/*
Plugin Name: GST Audit Exporter for WooCommerce
Description: GST Audit Exporter v1.3.2 – A comprehensive WooCommerce extension to simplify GST compliance and reporting for Indian e-commerce businesses. Developed by Satheesh Kumar S at Mallow Technologies Private Limited. Export detailed order data for GST filing, manage HSN codes, and generate Excel reports.
Version: 1.3.2
Requires at least: 5.0
Requires PHP: 7.2
Author: Satheesh Kumar S 
Author URI: https://github.com/pplcallmesatz/
Website: https://github.com/pplcallmesatz/
License: GPLv2 or later
License URI: https://www.gnu.org/licenses/gpl-2.0.html
Text Domain: gst-audit-exporter-for-woocommerce

Key Features:
- Detailed GST Reports: Export WooCommerce order data with HSN codes, tax breakdowns, and all essential fields for GST returns.
- Excel Export: Download monthly GST reports in Excel format, ready for upload or sharing with your accountant.
- HSN Code Management: Easily view, edit, and bulk update HSN codes for all products from a dedicated admin interface.
- Tax Class & Rate Breakdown: See a clear breakdown of all applicable tax classes and rates for each order item and shipping.
- Customer Details: Export includes customer name, city, and pincode for accurate record-keeping and compliance.
- User-Friendly Interface: Filter reports by month, customize pagination, and preview data before export.
- Seamless WooCommerce Integration: Works out-of-the-box with your existing WooCommerce setup.
*/

if ( ! defined( 'ABSPATH' ) ) exit; // Exit if accessed directly

class GST_Audit_Exporter {
    public function __construct() {
        add_action('admin_menu', [$this, 'register_menu']);
        add_action('admin_enqueue_scripts', [$this, 'enqueue_admin_assets']);
        add_action('wp_ajax_gst_audit_export_excel', [$this, 'ajax_export_excel']);
        // HSN code field in product
        add_action('woocommerce_product_options_general_product_data', [$this, 'add_hsn_code_field']);
        add_action('woocommerce_process_product_meta', [$this, 'save_hsn_code_field']);
        // HSN Checker AJAX
        add_action('wp_ajax_gst_audit_save_hsn_code', [$this, 'ajax_save_hsn_code']);
        add_action('wp_ajax_gst_audit_hsn_checker', [$this, 'ajax_hsn_checker']);
        // Mail settings AJAX
        add_action('wp_ajax_gst_audit_save_mail_settings', [$this, 'ajax_save_mail_settings']);
        add_action('wp_ajax_gst_audit_regenerate_key', [$this, 'ajax_regenerate_key']);
        // Cron scheduling for monthly emails
        add_action('gst_audit_monthly_email', [$this, 'send_monthly_gst_report']);
        add_action('init', [$this, 'schedule_monthly_email']);
        // Ensure cron is triggered
        add_action('wp', [$this, 'ensure_cron_triggered']);
        // Add custom hourly check for more reliable scheduling
        add_action('gst_audit_hourly_check', [$this, 'check_and_send_monthly_email']);
        add_action('init', [$this, 'schedule_hourly_check']);
        // Add more frequent check every 5 minutes for better reliability
        add_action('gst_audit_minute_check', [$this, 'check_and_send_monthly_email']);
        add_action('init', [$this, 'schedule_minute_check']);
        // Add custom cron interval for 5 minutes
        add_filter('cron_schedules', [$this, 'add_custom_cron_intervals']);
        // Create key management tables
        add_action('admin_init', [$this, 'create_key_management_tables']);
        // Add aggressive cron checking on every page load
        add_action('init', [$this, 'aggressive_cron_check'], 20);
        // Add public URL endpoint for triggering emails
        add_action('init', [$this, 'handle_public_email_trigger']);
    }

    public function enqueue_admin_assets($hook) {
        if ($hook !== 'toplevel_page_gst-audit-export') return;
        wp_enqueue_script('gst-audit-exporter-js', plugins_url('gst-audit-exporter.js', __FILE__), ['jquery'], '1.3.2', true);
        wp_localize_script('gst-audit-exporter-js', 'gst_audit_export', [
            'ajax_url' => admin_url('admin-ajax.php'),
            'nonce' => wp_create_nonce('gst_audit_export_excel'),
            'hsn_nonce' => wp_create_nonce('gst_audit_save_hsn_code'),
        ]);
    }


    public function register_menu() {
        add_menu_page(
            'GST Audit Export',
            'GST Audit Export',
            'manage_options',
            'gst-audit-export',
            [$this, 'render_admin_page'],
            'dashicons-media-spreadsheet',
            56
        );
    }

    public function render_admin_page() {
        if ( ! current_user_can( 'manage_options' ) ) {
            wp_die( esc_html( __( 'You do not have sufficient permissions to access this page.', 'gst-audit-exporter-for-woocommerce' ) ) );
        }
        $tab = isset($_GET['tab']) ? sanitize_text_field(wp_unslash($_GET['tab'])) : 'report';
        echo '<div class="wrap"><h1>GST Audit Export</h1>';
        echo '<h2 class="nav-tab-wrapper">';
        echo '<a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=report')) . '" class="nav-tab' . ($tab === 'report' ? ' nav-tab-active' : '') . '">Report</a>';
        echo '<a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=hsn')) . '" class="nav-tab' . ($tab === 'hsn' ? ' nav-tab-active' : '') . '">HSN Checker</a>';
        echo '<a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=mail')) . '" class="nav-tab' . ($tab === 'mail' ? ' nav-tab-active' : '') . '">Mail Settings</a>';
        echo '<a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=webhook')) . '" class="nav-tab' . ($tab === 'webhook' ? ' nav-tab-active' : '') . '">Webhook</a>';
        echo '</h2>';
        if ($tab === 'report') {
            // Verify nonce for POST requests
            if (isset($_POST['gst_audit_month']) || isset($_POST['gst_audit_per_page'])) {
                // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
                if (!isset($_POST['gst_audit_nonce']) || !wp_verify_nonce(wp_unslash($_POST['gst_audit_nonce']), 'gst_audit_filter')) {
                    wp_die(esc_html(__('Security check failed. Please try again.', 'gst-audit-exporter-for-woocommerce')));
                }
            }
            $selected_month = isset($_POST['gst_audit_month']) ? sanitize_text_field(wp_unslash($_POST['gst_audit_month'])) : (isset($_GET['gst_audit_month']) ? sanitize_text_field(wp_unslash($_GET['gst_audit_month'])) : gmdate('Y-m'));
            $per_page = isset($_POST['gst_audit_per_page']) ? max(10, intval(wp_unslash($_POST['gst_audit_per_page']))) : (isset($_GET['gst_audit_per_page']) ? max(10, intval(wp_unslash($_GET['gst_audit_per_page']))) : 20);
            // Month/per-page filter form (GET for pagination, POST for per-page change)
            echo '<form method="get" style="display:inline-block; margin-right:20px;">';
            echo '<input type="hidden" name="page" value="gst-audit-export" />';
            echo '<label for="gst_audit_month">Select Month: </label>';
            echo '<input type="month" id="gst_audit_month" name="gst_audit_month" value="' . esc_attr($selected_month) . '" />';
            echo '<input type="hidden" name="gst_audit_per_page" value="' . esc_attr($per_page) . '" />';
            echo '<input type="submit" class="button" value="Filter" style="margin-left:10px;" />';
            echo '</form>';
            // Download button (AJAX)
            echo '<button id="gst-audit-export-download-btn" class="button button-primary" style="margin-left:10px;">Download Excel</button>';
            echo '<span id="gst-audit-export-spinner" style="display:none; margin-left:10px;"><span class="spinner is-active" style="float:none;display:inline-block;"></span> Generating...</span>';
            echo '<hr />';
            echo '<div id="gst-audit-export-table">';
            $this->render_preview_table($selected_month);
            echo '</div>';
        } elseif ($tab === 'hsn') {
            echo '<div id="gst-audit-hsn-checker">';
            echo '<h2>HSN Checker</h2>';
            echo '<div id="gst-hsn-table-container">';
            $this->render_hsn_checker();
            echo '</div>';
            echo '<div id="gst-hsn-loader" style="display:none;text-align:center;margin:20px 0;"><span class="spinner is-active" style="float:none;display:inline-block;"></span> Loading...</div>';
            echo '</div>';
        } elseif ($tab === 'mail') {
            echo '<div id="gst-audit-mail-settings">';
            echo '<h2>Mail Settings</h2>';
            $this->render_mail_settings();
            echo '</div>';
        } elseif ($tab === 'webhook') {
            echo '<div id="gst-audit-webhook-settings">';
            echo '<h2>Webhook Settings</h2>';
            $this->render_webhook_settings();
            echo '</div>';
        }
        echo '</div>';
    }

    public function render_hsn_checker($paged = null, $per_page = null) {
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for pagination are safe
        if ($paged === null) $paged = isset($_GET['product_page']) ? max(1, intval(wp_unslash($_GET['product_page']))) : 1;
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for pagination are safe
        if ($per_page === null) $per_page = isset($_GET['hsn_per_page']) ? max(10, intval(wp_unslash($_GET['hsn_per_page']))) : 20;
        $args = [
            'post_type' => 'product',
            'posts_per_page' => $per_page,
            'paged' => $paged,
            'post_status' => 'publish',
        ];
        $query = new WP_Query($args);
        // Top controls
        $total_pages = $query->max_num_pages;
        $total_items = $query->found_posts;
        $this->render_hsn_pagination($paged, $total_pages, $total_items, $per_page, true);
        echo '<table class="widefat striped"><thead><tr>';
        echo '<th>Product ID</th><th>Product Name</th><th>SKU</th><th>HSN Code</th><th>Status</th>';
        echo '</tr></thead><tbody>';
        if ($query->have_posts()) {
            while ($query->have_posts()) {
                $query->the_post();
                $product_id = get_the_ID();
                $product = wc_get_product($product_id);
                $hsn_code = $this->get_hsn_code($product_id);
                $sku = $product ? $product->get_sku() : '';
                $product_edit_url = admin_url('post.php?post=' . $product_id . '&action=edit');
                echo '<tr>';
                echo '<td>' . esc_html($product_id) . '</td>';
                echo '<td><a href="' . esc_url($product_edit_url) . '" target="_blank">' . esc_html(get_the_title()) . '</a></td>';
                echo '<td>' . esc_html($sku) . '</td>';
                echo '<td><input type="text" class="gst-hsn-input" data-product-id="' . esc_attr($product_id) . '" value="' . esc_attr($hsn_code) . '" style="width:120px;" /></td>';
                echo '<td><span class="gst-hsn-status" id="gst-hsn-status-' . esc_attr($product_id) . '"></span></td>';
                echo '</tr>';
            }
            wp_reset_postdata();
        } else {
            echo '<tr><td colspan="5" style="text-align:center;">No products found.</td></tr>';
        }
        echo '</tbody></table>';
        // Bottom controls
        $this->render_hsn_pagination($paged, $total_pages, $total_items, $per_page, false);
        // JS for AJAX save
        echo '<script>window.gst_audit_hsn_nonce = "' . esc_js(wp_create_nonce('gst_audit_save_hsn_code')) . '";</script>';
        
    }

    public function render_hsn_pagination($paged, $total_pages, $total_items, $per_page, $top) {
        $select_id = $top ? 'gst-hsn-per-page-top' : 'gst-hsn-per-page-bottom';
        echo '<div style="text-align:center; margin:10px 0;">';
        echo '<form class="gst-hsn-per-page-form" data-top="' . ($top ? '1' : '0') . '" style="display:inline-block; margin-right:20px;">';
        echo '<label for="' . esc_attr($select_id) . '">Show </label>';
        echo '<select name="hsn_per_page" id="' . esc_attr($select_id) . '" class="gst-hsn-per-page-select">';
        $options = [10, 20, 50, 100, 200, 500, 750, 1000];
        foreach ($options as $opt) {
            $sel = ($per_page == $opt) ? 'selected' : '';
            echo '<option value="' . esc_attr($opt) . '" ' . esc_attr($sel) . '>' . esc_html($opt) . '</option>';
        }
        echo '</select> items per page';
        echo '</form>';
        echo '<strong>Total items: ' . esc_html($total_items) . '</strong> | ';
        echo 'Page ' . esc_html($paged) . ' of ' . esc_html($total_pages) . ' ';
        if ($total_pages > 1) {
            for ($i = 1; $i <= $total_pages; $i++) {
                $active = $i === $paged ? 'font-weight:bold;text-decoration:underline;' : '';
                echo '<a href="#" class="gst-hsn-page-link" data-page="' . esc_attr($i) . '" style="display:inline-block; margin:0 8px; padding:4px 10px; border-radius:4px; background:' . esc_attr($i === $paged ? '#dbeafe' : 'transparent') . '; color:' . esc_attr($i === $paged ? '#1d4ed8' : '#222') . '; ' . esc_attr($active) . '">' . esc_html($i) . '</a>';
            }
        }
        echo '</div>';
    }

    public function ajax_hsn_checker() {
        check_ajax_referer('gst_audit_export_excel', 'nonce');
        $paged = isset($_POST['page']) ? max(1, intval(wp_unslash($_POST['page']))) : 1;
        $per_page = isset($_POST['per_page']) ? max(10, intval(wp_unslash($_POST['per_page']))) : 20;
        ob_start();
        $this->render_hsn_checker($paged, $per_page);
        $html = ob_get_clean();
        wp_send_json_success(['html' => $html]);
    }

    public function ajax_save_hsn_code() {
        check_ajax_referer('gst_audit_save_hsn_code', 'nonce');
        $product_id = isset($_POST['product_id']) ? intval(wp_unslash($_POST['product_id'])) : 0;
        $hsn_code = isset($_POST['hsn_code']) ? sanitize_text_field(wp_unslash($_POST['hsn_code'])) : '';
        if (!$product_id || get_post_type($product_id) !== 'product') {
            wp_send_json_error(['message' => 'Invalid product.']);
        }
        update_post_meta($product_id, 'hsn_code', $hsn_code);
        wp_send_json_success(['message' => 'HSN code saved.']);
    }

    private function get_tax_classes_and_slugs() {
        $tax_classes = array('Standard' => '');
        $wc_tax_classes = WC_Tax::get_tax_classes();
        foreach ($wc_tax_classes as $class) {
            $tax_classes[$class] = sanitize_title($class);
        }
        return $tax_classes;
    }

    private function get_tax_classes_and_rates() {
        // Check cache first
        $cache_key = 'gst_audit_tax_classes_and_rates';
        $cached_rates = wp_cache_get($cache_key, 'gst_audit');
        if (false !== $cached_rates) {
            return $cached_rates;
        }
        
        $tax_classes = array('Standard' => '');
        $wc_tax_classes = WC_Tax::get_tax_classes();
        foreach ($wc_tax_classes as $class) {
            $tax_classes[$class] = sanitize_title($class);
        }
        
        // For each class, get all rates using WooCommerce API
        $class_rates = [];
        foreach ($tax_classes as $class_name => $slug) {
            $rates = WC_Tax::get_rates_for_tax_class($slug);
            $class_rates[$class_name] = [];
            foreach ($rates as $rate_id => $rate) {
                // Handle both array and object formats
                $label = is_array($rate) ? $rate['label'] : $rate->label;
                $percent = is_array($rate) ? $rate['rate'] : $rate->rate;
                $class_rates[$class_name][] = [
                    'id' => $rate_id,
                    'label' => $label,
                    'percent' => $percent
                ];
            }
        }
        
        // Cache the results for 1 hour
        wp_cache_set($cache_key, $class_rates, 'gst_audit', HOUR_IN_SECONDS);
        
        return $class_rates;
    }

    private function render_pagination_controls($current_page, $total_pages, $total_items, $per_page, $selected_month) {
        $base_url = remove_query_arg(['gst_audit_page', 'gst_audit_per_page']);
        echo '<div style="text-align:center; margin: 10px 0;">';
            echo '<form method="post" style="display:inline-block; margin-right:20px;">';
            wp_nonce_field('gst_audit_filter', 'gst_audit_nonce');
            echo '<input type="hidden" name="gst_audit_month" value="' . esc_attr($selected_month) . '" />';
        echo '<label for="gst_audit_per_page">Show </label>';
        echo '<select name="gst_audit_per_page" id="gst_audit_per_page" onchange="this.form.submit()">';
        $options = [10, 20, 50, 100, 200, 500, 750, 1000];
        foreach ($options as $opt) {
            $sel = ($per_page == $opt) ? 'selected' : '';
            echo '<option value="' . esc_attr($opt) . '" ' . esc_attr($sel) . '>' . esc_html($opt) . '</option>';
        }
        echo '</select> items per page';
        echo '</form>';
        echo '<strong>Total items: ' . esc_html($total_items) . '</strong> | ';
        echo 'Page ' . esc_html($current_page) . ' of ' . esc_html($total_pages) . ' ';
        // Pagination links
        if ($current_page > 1) {
            echo '<a style="margin:0 10px;" href="' . esc_url(add_query_arg(['gst_audit_page' => $current_page - 1, 'gst_audit_per_page' => $per_page], $base_url)) . '">&laquo; Prev</a>';
        }
        if ($current_page < $total_pages) {
            echo '<a style="margin:0 10px;" href="' . esc_url(add_query_arg(['gst_audit_page' => $current_page + 1, 'gst_audit_per_page' => $per_page], $base_url)) . '">Next &raquo;</a>';
        }
        echo '</div>';
    }

    private function render_preview_table($selected_month) {
        global $wpdb;
        $default_per_page = 20;
        // Verify nonce for POST requests
        if (isset($_POST['gst_audit_per_page'])) {
            // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
            if (!isset($_POST['gst_audit_nonce']) || !wp_verify_nonce(wp_unslash($_POST['gst_audit_nonce']), 'gst_audit_filter')) {
                wp_die(esc_html(__('Security check failed. Please try again.', 'gst-audit-exporter-for-woocommerce')));
            }
        }
        $per_page = isset($_POST['gst_audit_per_page']) ? max(10, intval(wp_unslash($_POST['gst_audit_per_page']))) : (isset($_GET['gst_audit_per_page']) ? max(10, intval(wp_unslash($_GET['gst_audit_per_page']))) : $default_per_page);
        $current_page = isset($_GET['gst_audit_page']) ? max(1, intval(wp_unslash($_GET['gst_audit_page']))) : 1;
        $start_date = $selected_month . '-01 00:00:00';
        $end_date = gmdate('Y-m-t 23:59:59', strtotime($start_date));
        $args = array(
            'status' => array('wc-completed', 'wc-processing'),
            'type'   => 'shop_order',
            'date_query' => array(
                array(
                    'after'     => $start_date,
                    'before'    => $end_date,
                    'inclusive' => true,
                ),
            ),
            'limit' => -1, // Get all for counting
        );
        $all_orders = wc_get_orders($args);
        $total_items = 0;
        $order_items = [];
        foreach ($all_orders as $order) {
            foreach ($order->get_items() as $item) {
                $order_items[] = ['order' => $order, 'item' => $item];
            }
        }
        $total_items = count($order_items);
        $total_pages = max(1, ceil($total_items / $per_page));
        $current_page = min($current_page, $total_pages);
        $offset = ($current_page - 1) * $per_page;
        $order_items_page = array_slice($order_items, $offset, $per_page);
        $class_rates = $this->get_tax_classes_and_rates();
        // Top pagination
        $this->render_pagination_controls($current_page, $total_pages, $total_items, $per_page, $selected_month);
        // Build header
        // Add scrollable wrapper
        echo '<div style="overflow-x:auto; width:100%;">';
        echo '<table class="widefat striped">';
        echo '<thead>';
        // First row: parent columns
        echo '<tr>';
        echo '<th rowspan="2" style="text-align:center; vertical-align:middle;">Order Date</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Order ID</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Invoice Number</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Order Status</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Name</th><th rowspan="2" style="text-align:center; vertical-align:middle;">City</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Pincode</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Product Name</th><th rowspan="2" style="text-align:center; vertical-align:middle;">HSN Code</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Price (Inc Tax)</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Qty</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Total (Inc Tax)</th><th rowspan="2" style="text-align:center; vertical-align:middle;">Total Price (Excl Tax)</th>';
        foreach ($class_rates as $class_name => $rates) {
            $colspan = max(1, count($rates));
            echo '<th colspan="' . esc_attr($colspan) . '">' . esc_html($class_name) . '</th>';
        }
        echo '</tr>';
        // Second row: tax rate columns
        echo '<tr>';
        foreach ($class_rates as $rates) {
            if (empty($rates)) {
                echo '<th>-</th>';
            } else {
                foreach ($rates as $rate) {
                    $label = $rate['label'] ? $rate['label'] : 'Rate';
                    $percent = $rate['percent'] !== '' ? ' (' . $rate['percent'] . '%)' : '';
                    echo '<th>' . esc_html($label . $percent) . '</th>';
                }
            }
        }
        echo '</tr>';
        echo '</thead>';
        echo '<tbody>';
        $total_tax_cols = 0;
        foreach ($class_rates as $rates) {
            $total_tax_cols += max(1, count($rates));
        }
        if (empty($order_items_page)) {
            echo '<tr><td colspan="' . esc_attr(9 + $total_tax_cols) . '" style="text-align:center;">No orders found for this month.</td></tr>';
        } else {
            // Group by order for shipping row
            $orders_seen = [];
            foreach ($order_items_page as $row) {
                $order = $row['order'];
                $item = $row['item'];
                $order_id = $order->get_id();
                $order_date = $order->get_date_created() ? $order->get_date_created()->date('Y-m-d') : '';
                $invoice_number = $order->get_meta('_wcpdf_invoice_number');
                $order_status = $order->get_status();
                $order_name = $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
                $order_city = $order->get_billing_city();
                $order_pincode = $order->get_billing_postcode();
                $product = $item->get_product();
                $product_name = $item->get_name();
                // For variations, get the parent product ID for HSN code
                $variation_id = $item->get_variation_id();
                $product_id = $product ? $product->get_id() : 0;
                
                // If this is a variation, get the parent product ID
                if ($variation_id && $variation_id > 0) {
                    $variation_product = wc_get_product($variation_id);
                    if ($variation_product && $variation_product->get_parent_id()) {
                        $product_id = $variation_product->get_parent_id();
                    }
                }
                
                $hsn_code = $product ? $this->get_hsn_code($product_id, $variation_id) : '';
                $price_inc_tax = wc_format_decimal($order->get_item_total($item, true, false), 2);
                $qty = $item->get_quantity();
                $total_excl_tax = wc_format_decimal($order->get_line_total($item, false, false), 2);
                // Prepare tax rate columns
                $tax_amounts = [];
                foreach ($class_rates as $class_name => $rates) {
                    if (empty($rates)) {
                        $tax_amounts[$class_name] = ['-'];
                    } else {
                        foreach ($rates as $rate) {
                            $tax_amounts[$class_name][$rate['id']] = '';
                        }
                    }
                }
                $item_taxes = $item->get_taxes();
                if (!empty($item_taxes['total'])) {
                    foreach ($item_taxes['total'] as $tax_rate_id => $tax_amount) {
                        foreach ($class_rates as $class_name => $rates) {
                            foreach ($rates as $rate) {
                                if ($rate['id'] == $tax_rate_id) {
                                    $tax_amounts[$class_name][$tax_rate_id] = wc_format_decimal($tax_amount, 2);
                                }
                            }
                        }
                    }
                }
                echo '<tr>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_date) . '</td>';
                $order_edit_url = admin_url('post.php?post=' . $order_id . '&action=edit');
                echo '<td style="text-align:center; vertical-align:middle;"><a href="' . esc_url($order_edit_url) . '" target="_blank">' . esc_html($order_id) . '</a></td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($invoice_number) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_status) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_name) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_city) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_pincode) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($product_name) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($hsn_code) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($price_inc_tax) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($qty) . '</td>';
                $line_total_inc_tax = wc_format_decimal($price_inc_tax * $qty, 2);
                echo '<td style="background:#d1fae5; font-weight:bold; text-align:center; vertical-align:middle;">' . esc_html($line_total_inc_tax) . '</td>';
                echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($total_excl_tax) . '</td>';
                foreach ($class_rates as $class_name => $rates) {
                    if (empty($rates)) {
                        echo '<td>-</td>';
                    } else {
                        foreach ($rates as $rate) {
                            $val = isset($tax_amounts[$class_name][$rate['id']]) ? $tax_amounts[$class_name][$rate['id']] : '';
                            echo '<td>' . esc_html($val) . '</td>';
                        }
                    }
                }
                echo '</tr>';
                // Only add shipping row once per order
                if (!isset($orders_seen[$order_id])) {
                    $orders_seen[$order_id] = true;
                    // Shipping row (one per order, totals)
                    $shipping_total = 0;
                    $shipping_total_tax = 0;
                    $shipping_tax_amounts = [];
                    foreach ($class_rates as $class_name => $rates) {
                        if (empty($rates)) {
                            $shipping_tax_amounts[$class_name] = ['-'];
                        } else {
                            foreach ($rates as $rate) {
                                $shipping_tax_amounts[$class_name][$rate['id']] = '';
                            }
                        }
                    }
                    foreach ($order->get_items('shipping') as $shipping_item) {
                        $shipping_total += (float) $shipping_item->get_total();
                        $shipping_total_tax += (float) $shipping_item->get_total_tax();
                        $shipping_taxes = $shipping_item->get_taxes();
                        if (!empty($shipping_taxes['total'])) {
                            foreach ($shipping_taxes['total'] as $tax_rate_id => $tax_amount) {
                                foreach ($class_rates as $class_name => $rates) {
                                    foreach ($rates as $rate) {
                                        if ($rate['id'] == $tax_rate_id) {
                                            $prev = isset($shipping_tax_amounts[$class_name][$tax_rate_id]) && $shipping_tax_amounts[$class_name][$tax_rate_id] !== '' ? (float)$shipping_tax_amounts[$class_name][$tax_rate_id] : 0;
                                            $shipping_tax_amounts[$class_name][$tax_rate_id] = wc_format_decimal($prev + (float)$tax_amount, 2);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if ($shipping_total > 0 || $shipping_total_tax > 0) {
                        echo '<tr>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_date) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;"><a href="' . esc_url($order_edit_url) . '" target="_blank">' . esc_html($order_id) . '</a></td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($invoice_number) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_status) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_name) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_city) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html($order_pincode) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">Shipping</td>';
                        echo '<td style="text-align:center; vertical-align:middle;"></td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html(wc_format_decimal($shipping_total + $shipping_total_tax, 2)) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">1</td>';
                        echo '<td style="background:#d1fae5; font-weight:bold; text-align:center; vertical-align:middle;">' . esc_html(wc_format_decimal($shipping_total + $shipping_total_tax, 2)) . '</td>';
                        echo '<td style="text-align:center; vertical-align:middle;">' . esc_html(wc_format_decimal($shipping_total, 2)) . '</td>';
                        foreach ($class_rates as $class_name => $rates) {
                            if (empty($rates)) {
                                echo '<td>-</td>';
                            } else {
                                foreach ($rates as $rate) {
                                    $val = isset($shipping_tax_amounts[$class_name][$rate['id']]) ? $shipping_tax_amounts[$class_name][$rate['id']] : '';
                                    echo '<td>' . esc_html($val) . '</td>';
                                }
                            }
                        }
                        echo '</tr>';
                    }
                }
            }
        }
        echo '</tbody></table>';
        echo '</div>';
        // Bottom pagination
        $this->render_pagination_controls($current_page, $total_pages, $total_items, $per_page, $selected_month);
    }

    public function ajax_export_excel() {
        check_ajax_referer('gst_audit_export_excel');
        $selected_month = isset($_POST['gst_audit_month']) ? sanitize_text_field(wp_unslash($_POST['gst_audit_month'])) : gmdate('Y-m');
        $this->handle_excel_export($selected_month, true);
    }

    private function handle_excel_export($selected_month, $is_ajax = false) {
        if ((isset($_POST['gst_audit_export_excel']) && check_admin_referer('gst_audit_export_excel', 'gst_audit_export_excel_nonce')) || $is_ajax) {
            require_once __DIR__ . '/vendor/autoload.php';
            if (!class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet')) {
                if ($is_ajax) {
                    wp_send_json_error(['message' => 'PhpSpreadsheet library not found. Please install it in the plugin folder for Excel export to work.']);
                } else {
                    add_action('admin_notices', function() {
                        echo '<div class="notice notice-error"><p>PhpSpreadsheet library not found. Please install it in the plugin folder for Excel export to work.</p></div>';
                    });
                }
                return;
            }
            global $wpdb;
            $start_date = $selected_month . '-01 00:00:00';
            $end_date = gmdate('Y-m-t 23:59:59', strtotime($start_date));
            $args = array(
                'status' => array('wc-completed', 'wc-processing'),
                'type'   => 'shop_order',
                'date_query' => array(
                    array(
                        'after'     => $start_date,
                        'before'    => $end_date,
                        'inclusive' => true,
                    ),
                ),
                'limit' => -1, // All orders for export
            );
            $all_orders = wc_get_orders($args);
            $class_rates = $this->get_tax_classes_and_rates();
            // Build header rows
            $header1 = ['Order Date','Order ID','Invoice Number','Order Status','Name','City','Pincode','Product Name','HSN Code','Price (Inc Tax)','Qty','Total (Inc Tax)','Total Price (Excl Tax)'];
            $header2 = ['', '', '', '', '', '', '', '', '', '', '', '', ''];
            foreach ($class_rates as $class_name => $rates) {
                $colspan = max(1, count($rates));
                $header1[] = $class_name;
                for ($i = 1; $i < $colspan; $i++) {
                    $header1[] = '';
                }
                if (empty($rates)) {
                    $header2[] = '-';
                } else {
                    foreach ($rates as $rate) {
                        $label = $rate['label'] ? $rate['label'] : 'Rate';
                        $percent = $rate['percent'] !== '' ? ' (' . $rate['percent'] . '%)' : '';
                        $header2[] = $label . $percent;
                    }
                }
            }
            $rows = [];
            $rows[] = $header1;
            $rows[] = $header2;
            foreach ($all_orders as $order) {
                $order_date = $order->get_date_created() ? $order->get_date_created()->date('Y-m-d') : '';
                $order_id = $order->get_id();
                $invoice_number = $order->get_meta('_wcpdf_invoice_number');
                $order_status = $order->get_status();
                $order_name = $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
                $order_city = $order->get_billing_city();
                $order_pincode = $order->get_billing_postcode();
                // Product rows
                foreach ($order->get_items() as $item) {
                    $product = $item->get_product();
                    $product_name = $item->get_name();
                    // For variations, get the parent product ID for HSN code
                $variation_id = $item->get_variation_id();
                $product_id = $product ? $product->get_id() : 0;
                
                // If this is a variation, get the parent product ID
                if ($variation_id && $variation_id > 0) {
                    $variation_product = wc_get_product($variation_id);
                    if ($variation_product && $variation_product->get_parent_id()) {
                        $product_id = $variation_product->get_parent_id();
                    }
                }
                
                $hsn_code = $product ? $this->get_hsn_code($product_id, $variation_id) : '';
                    $price_inc_tax = wc_format_decimal($order->get_item_total($item, true, false), 2);
                    $qty = $item->get_quantity();
                    $total_excl_tax = wc_format_decimal($order->get_line_total($item, false, false), 2);
                    $line_total_inc_tax = wc_format_decimal($price_inc_tax * $qty, 2);
                    $tax_amounts = [];
                    foreach ($class_rates as $class_name => $rates) {
                        if (empty($rates)) {
                            $tax_amounts[$class_name] = ['-'];
                        } else {
                            foreach ($rates as $rate) {
                                $tax_amounts[$class_name][$rate['id']] = '';
                            }
                        }
                    }
                    $item_taxes = $item->get_taxes();
                    if (!empty($item_taxes['total'])) {
                        foreach ($item_taxes['total'] as $tax_rate_id => $tax_amount) {
                            foreach ($class_rates as $class_name => $rates) {
                                foreach ($rates as $rate) {
                                    if ($rate['id'] == $tax_rate_id) {
                                        $tax_amounts[$class_name][$tax_rate_id] = wc_format_decimal($tax_amount, 2);
                                    }
                                }
                            }
                        }
                    }
                    $row = [$order_date, $order_id, $invoice_number, $order_status, $order_name, $order_city, $order_pincode, $product_name, $hsn_code, $price_inc_tax, $qty, $line_total_inc_tax, $total_excl_tax];
                    foreach ($class_rates as $class_name => $rates) {
                        if (empty($rates)) {
                            $row[] = '-';
                        } else {
                            foreach ($rates as $rate) {
                                $row[] = isset($tax_amounts[$class_name][$rate['id']]) ? $tax_amounts[$class_name][$rate['id']] : '';
                            }
                        }
                    }
                    $rows[] = $row;
                }
                // Shipping row (one per order, totals)
                $shipping_total = 0;
                $shipping_total_tax = 0;
                $shipping_tax_amounts = [];
                foreach ($class_rates as $class_name => $rates) {
                    if (empty($rates)) {
                        $shipping_tax_amounts[$class_name] = ['-'];
                    } else {
                        foreach ($rates as $rate) {
                            $shipping_tax_amounts[$class_name][$rate['id']] = '';
                        }
                    }
                }
                foreach ($order->get_items('shipping') as $shipping_item) {
                    $shipping_total += (float) $shipping_item->get_total();
                    $shipping_total_tax += (float) $shipping_item->get_total_tax();
                    $shipping_taxes = $shipping_item->get_taxes();
                    if (!empty($shipping_taxes['total'])) {
                        foreach ($shipping_taxes['total'] as $tax_rate_id => $tax_amount) {
                            foreach ($class_rates as $class_name => $rates) {
                                foreach ($rates as $rate) {
                                    if ($rate['id'] == $tax_rate_id) {
                                        $prev = isset($shipping_tax_amounts[$class_name][$tax_rate_id]) && $shipping_tax_amounts[$class_name][$tax_rate_id] !== '' ? (float)$shipping_tax_amounts[$class_name][$tax_rate_id] : 0;
                                        $shipping_tax_amounts[$class_name][$tax_rate_id] = wc_format_decimal($prev + (float)$tax_amount, 2);
                                    }
                                }
                            }
                        }
                    }
                }
                if ($shipping_total > 0 || $shipping_total_tax > 0) {
                    $row = [
                        $order_date,
                        $order_id,
                        $invoice_number,
                        $order_status,
                        $order_name,
                        $order_city,
                        $order_pincode,
                        'Shipping',
                        '', // HSN Code blank
                        wc_format_decimal($shipping_total + $shipping_total_tax, 2), // Price (Inc Tax)
                        1, // Qty
                        wc_format_decimal($shipping_total + $shipping_total_tax, 2), // Total (Inc Tax)
                        wc_format_decimal($shipping_total, 2) // Total Price (Excl Tax)
                    ];
                    foreach ($class_rates as $class_name => $rates) {
                        if (empty($rates)) {
                            $row[] = '-';
                        } else {
                            foreach ($rates as $rate) {
                                $row[] = isset($shipping_tax_amounts[$class_name][$rate['id']]) ? $shipping_tax_amounts[$class_name][$rate['id']] : '';
                            }
                        }
                    }
                    $rows[] = $row;
                }
            }
            if (ob_get_length()) ob_end_clean();
            $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            // Styling: bold headers, background color, borders, auto width
            $headerRowCount = 2;
            $colCount = count($header1);
            // Set header styles
            $headerStyle = [
                'font' => ['bold' => true],
                'fill' => [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'startColor' => ['rgb' => 'D9E1F2']
                ],
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['rgb' => 'B4B4B4']
                    ]
                ]
            ];
            $sheet->getStyle('A1:' . $sheet->getCellByColumnAndRow($colCount, $headerRowCount)->getCoordinate())
                ->applyFromArray($headerStyle);
            // Set row height for headers
            for ($i = 1; $i <= $headerRowCount; $i++) {
                $sheet->getRowDimension($i)->setRowHeight(24);
            }
            // Write data rows
            foreach ($rows as $rowIdx => $row) {
                foreach ($row as $colIdx => $value) {
                    $cell = $sheet->getCellByColumnAndRow($colIdx + 1, $rowIdx + 1);
                    $cell->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                }
            }
            // Set borders and row height for all data rows
            $lastRow = count($rows);
            $sheet->getStyle('A1:' . $sheet->getCellByColumnAndRow($colCount, $lastRow)->getCoordinate())
                ->getBorders()->getAllBorders()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color('B4B4B4'));
            for ($i = $headerRowCount + 1; $i <= $lastRow; $i++) {
                $sheet->getRowDimension($i)->setRowHeight(20);
            }
            // Auto-size columns
            for ($col = 1; $col <= $colCount; $col++) {
                $sheet->getColumnDimensionByColumn($col)->setAutoSize(true);
            }
            $filename = 'gst-audit-export-' . $selected_month . '.xlsx';
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $filename . '"');
            header('Cache-Control: max-age=0');
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
            $writer->save('php://output');
            exit;
        }
    }

    public function add_hsn_code_field() {
        woocommerce_wp_text_input([
            'id' => 'hsn_code',
            'label' => __('HSN Code', 'gst-audit-exporter-for-woocommerce'),
            'desc_tip' => true,
            'description' => __('Enter the HSN code for GST reporting.', 'gst-audit-exporter-for-woocommerce'),
            'type' => 'text',
        ]);
    }

    public function save_hsn_code_field($post_id) {
        // Verify nonce for security
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (!isset($_POST['woocommerce_meta_nonce']) || !wp_verify_nonce(wp_unslash($_POST['woocommerce_meta_nonce']), 'woocommerce_save_data')) {
            return;
        }
        $hsn_code = isset($_POST['hsn_code']) ? sanitize_text_field(wp_unslash($_POST['hsn_code'])) : '';
        update_post_meta($post_id, 'hsn_code', $hsn_code);
    }

    public function render_mail_settings() {
        // Handle form submission first
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (isset($_POST['save_mail_settings']) && isset($_POST['gst_audit_mail_nonce']) && wp_verify_nonce(wp_unslash($_POST['gst_audit_mail_nonce']), 'gst_audit_mail_settings')) {
            $this->save_mail_settings();
        }
        
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (isset($_POST['send_test_mail']) && isset($_POST['gst_audit_mail_nonce']) && wp_verify_nonce(wp_unslash($_POST['gst_audit_mail_nonce']), 'gst_audit_mail_settings')) {
            $this->send_test_mail();
        }
        
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (isset($_POST['trigger_manual_email']) && isset($_POST['gst_audit_mail_nonce']) && wp_verify_nonce(wp_unslash($_POST['gst_audit_mail_nonce']), 'gst_audit_mail_settings')) {
            $this->trigger_manual_email();
        }
        
        // Get form data after processing form submission
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        $mail_recipient = get_option('gst_audit_mail_recipient', '');
        $mail_day = get_option('gst_audit_mail_day', '1');
        $mail_time = get_option('gst_audit_mail_time', '09:00');
        
        echo '<form id="gst-mail-settings-form" method="post">';
        wp_nonce_field('gst_audit_mail_settings', 'gst_audit_mail_nonce');
        
        echo '<table class="form-table">';
        echo '<tr>';
        echo '<th scope="row"><label for="mail_enabled">Enable Automatic Email</label></th>';
        echo '<td>';
        echo '<label class="gst-toggle">';
        echo '<input type="checkbox" id="mail_enabled" name="mail_enabled" value="yes"' . checked($mail_enabled, 'yes', false) . ' />';
        echo '<span class="gst-toggle-slider"></span>';
        echo '</label>';
        echo '<p class="description">Send monthly GST reports automatically via email</p>';
        echo '</td>';
        echo '</tr>';
        
        echo '<tr>';
        echo '<th scope="row"><label for="mail_recipient">Email Recipients</label></th>';
        echo '<td>';
        echo '<input type="text" id="mail_recipient" name="mail_recipient" value="' . esc_attr($mail_recipient) . '" class="regular-text" required />';
        echo '<p class="description">Enter email addresses separated by commas (e.g., admin@example.com, manager@example.com)</p>';
        echo '</td>';
        echo '</tr>';
        
        echo '<tr>';
        echo '<th scope="row"><label for="mail_day">Send on Day</label></th>';
        echo '<td>';
        echo '<select id="mail_day" name="mail_day">';
        for ($i = 1; $i <= 28; $i++) {
            $selected = selected($mail_day, $i, false);
            echo '<option value="' . esc_attr($i) . '"' . esc_attr($selected) . '>' . esc_html($i) . '</option>';
        }
        echo '</select>';
        echo '<p class="description">Day of each month to send the report (1-28)</p>';
        echo '</td>';
        echo '</tr>';
        
        echo '<tr>';
        echo '<th scope="row"><label for="mail_time">Send at Time</label></th>';
        echo '<td>';
        echo '<input type="time" id="mail_time" name="mail_time" value="' . esc_attr($mail_time) . '" />';
        echo '<p class="description">Time of day to send the report (24-hour format)</p>';
        echo '</td>';
        echo '</tr>';
        echo '</table>';
        
        echo '<p class="submit">';
        echo '<input type="submit" name="save_mail_settings" class="button-primary" value="Save Settings" />';
        echo '<input type="submit" name="send_test_mail" class="button-secondary" value="Send Test Mail" style="margin-left: 10px;" onclick="return confirm(\'This will send a test email with the previous month\\\'s GST report. Continue?\');" />';
        echo '</p>';
        echo '</form>';
        
        // Show detailed cron status
        $next_scheduled = wp_next_scheduled('gst_audit_monthly_email');
        $hourly_scheduled = wp_next_scheduled('gst_audit_hourly_check');
        $minute_scheduled = wp_next_scheduled('gst_audit_minute_check');
        $cron_disabled = defined('DISABLE_WP_CRON') && DISABLE_WP_CRON;
        $cron_array = _get_cron_array();
        $last_sent_month = get_option('gst_audit_last_sent_month', '');
        
        
        // Mail settings content only - key management moved to webhook tab
        
        // Add manual trigger for testing
        if ($mail_enabled === 'yes' && !empty($mail_recipient)) {
            echo '<form method="post" style="margin-top: 20px;">';
            wp_nonce_field('gst_audit_mail_settings', 'gst_audit_mail_nonce');
            echo '<input type="submit" name="trigger_manual_email" class="button-secondary" value="Trigger Manual Email Now" onclick="return confirm(\'This will send the monthly email immediately. Continue?\');" />';
            echo '<p class="description">Use "Trigger Manual Email Now" to test email delivery.</p>';
            echo '</form>';
        }
        
        // Form submission is now handled in render_mail_settings()
        
        // Add CSS for toggle button
        echo '<style>';
        echo '.gst-toggle { position: relative; display: inline-block; width: 60px; height: 34px; }';
        echo '.gst-toggle input { opacity: 0; width: 0; height: 0; }';
        echo '.gst-toggle-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 34px; }';
        echo '.gst-toggle-slider:before { position: absolute; content: ""; height: 26px; width: 26px; left: 4px; bottom: 4px; background-color: white; transition: .4s; border-radius: 50%; }';
        echo '.gst-toggle input:checked + .gst-toggle-slider { background-color: #2196F3; }';
        echo '.gst-toggle input:checked + .gst-toggle-slider:before { transform: translateX(26px); }';
        echo '</style>';
    }

    private function validate_and_sanitize_emails($email_string) {
        if (empty($email_string)) {
            return '';
        }
        
        // Split by comma and clean each email
        $emails = array_map('trim', explode(',', $email_string));
        $valid_emails = [];
        
        foreach ($emails as $email) {
            $email = sanitize_email($email);
            if (is_email($email)) {
                $valid_emails[] = $email;
            }
        }
        
        return implode(', ', $valid_emails);
    }

    public function create_key_management_tables() {
        global $wpdb;
        
        $table_name_keys = $wpdb->prefix . 'gst_audit_keys';
        $table_name_logs = $wpdb->prefix . 'gst_audit_key_logs';
        
        $charset_collate = $wpdb->get_charset_collate();
        
        // Create keys table
        $sql_keys = "CREATE TABLE IF NOT EXISTS $table_name_keys (
            id int(11) NOT NULL AUTO_INCREMENT,
            key_value varchar(64) NOT NULL,
            is_active tinyint(1) DEFAULT 1,
            created_at datetime DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            UNIQUE KEY key_value (key_value)
        ) $charset_collate;";
        
        // Create logs table
        $sql_logs = "CREATE TABLE IF NOT EXISTS $table_name_logs (
            id int(11) NOT NULL AUTO_INCREMENT,
            key_id int(11) NOT NULL,
            access_time datetime DEFAULT CURRENT_TIMESTAMP,
            ip_address varchar(45) NOT NULL,
            user_agent text,
            browser varchar(100),
            os varchar(100),
            location varchar(255),
            success tinyint(1) DEFAULT 1,
            error_message text,
            PRIMARY KEY (id),
            KEY key_id (key_id),
            KEY access_time (access_time)
        ) $charset_collate;";
        
        require_once(ABSPATH . 'wp-admin/includes/upgrade.php');
        dbDelta($sql_keys);
        dbDelta($sql_logs);
    }

    public function ajax_regenerate_key() {
        check_ajax_referer('gst_audit_mail_settings', 'nonce');
        
        if (!current_user_can('manage_options')) {
            wp_send_json_error('Insufficient permissions');
        }
        
        global $wpdb;
        $table_name = $wpdb->prefix . 'gst_audit_keys';
        
        // Deactivate current active key
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        $wpdb->update(
            $table_name,
            ['is_active' => 0],
            ['is_active' => 1],
            ['%d'],
            ['%d']
        );
        
        // Clear cache after key deactivation
        wp_cache_delete('gst_audit_key_history', 'gst_audit');
        
        // Generate new key
        $new_key = wp_generate_password(32, false);
        
        // Insert new key
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        $wpdb->insert(
            $table_name,
            [
                'key_value' => $new_key,
                'is_active' => 1,
                'created_at' => current_time('mysql')
            ],
            ['%s', '%d', '%s']
        );
        
        // Clear cache after key creation
        wp_cache_delete('gst_audit_key_history', 'gst_audit');
        
        // Update option
        update_option('gst_audit_email_secret_key', $new_key);
        
        wp_send_json_success([
            'message' => 'Key regenerated successfully',
            'new_key' => $new_key
        ]);
    }

    private function log_key_access($key_value, $success = true, $error_message = '') {
        global $wpdb;
        
        $keys_table = $wpdb->prefix . 'gst_audit_keys';
        $logs_table = $wpdb->prefix . 'gst_audit_key_logs';
        
        // Get key ID
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.InterpolatedNotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQLPlaceholders.UnfinishedPrepare - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQLPlaceholders.MissingPlaceholder - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQLPlaceholders.UnknownPlaceholder - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.InterpolatedNotPrepared - Table name cannot be parameterized
        $key_id = $wpdb->get_var($wpdb->prepare(
            "SELECT id FROM " . $keys_table . " WHERE key_value = %s", // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared
            $key_value
        ));
        
        if (!$key_id) {
            return;
        }
        
        // Get browser and OS info
        $user_agent = isset($_SERVER['HTTP_USER_AGENT']) ? sanitize_text_field(wp_unslash($_SERVER['HTTP_USER_AGENT'])) : '';
        $browser = $this->get_browser_info($user_agent);
        $os = $this->get_os_info($user_agent);
        $location = $this->get_location_info();
        
        // Insert log
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        $wpdb->insert(
            $logs_table,
            [
                'key_id' => $key_id,
                'access_time' => current_time('mysql'),
                'ip_address' => $this->get_client_ip(),
                'user_agent' => $user_agent,
                'browser' => $browser,
                'os' => $os,
                'location' => $location,
                'success' => $success ? 1 : 0,
                'error_message' => $error_message
            ],
            ['%d', '%s', '%s', '%s', '%s', '%s', '%s', '%d', '%s']
        );
    }

    private function get_browser_info($user_agent) {
        if (strpos($user_agent, 'Chrome') !== false) return 'Chrome';
        if (strpos($user_agent, 'Firefox') !== false) return 'Firefox';
        if (strpos($user_agent, 'Safari') !== false) return 'Safari';
        if (strpos($user_agent, 'Edge') !== false) return 'Edge';
        if (strpos($user_agent, 'Opera') !== false) return 'Opera';
        return 'Unknown';
    }

    private function get_os_info($user_agent) {
        if (strpos($user_agent, 'Windows') !== false) return 'Windows';
        if (strpos($user_agent, 'Mac') !== false) return 'macOS';
        if (strpos($user_agent, 'Linux') !== false) return 'Linux';
        if (strpos($user_agent, 'Android') !== false) return 'Android';
        if (strpos($user_agent, 'iOS') !== false) return 'iOS';
        return 'Unknown';
    }

    private function get_location_info() {
        $ip = $this->get_client_ip();
        if ($ip === '127.0.0.1' || $ip === '::1') {
            return 'Localhost';
        }
        
        // Simple IP geolocation (you can integrate with a service like ipapi.co)
        $response = wp_remote_get("http://ip-api.com/json/{$ip}?fields=country,regionName,city");
        if (!is_wp_error($response)) {
            $data = json_decode(wp_remote_retrieve_body($response), true);
            if ($data && !isset($data['error'])) {
                return $data['city'] . ', ' . $data['regionName'] . ', ' . $data['country'];
            }
        }
        
        return 'Unknown';
    }

    private function get_client_ip() {
        $ip_keys = ['HTTP_CLIENT_IP', 'HTTP_X_FORWARDED_FOR', 'REMOTE_ADDR'];
        foreach ($ip_keys as $key) {
            if (array_key_exists($key, $_SERVER) === true) {
                foreach (explode(',', sanitize_text_field(wp_unslash($_SERVER[$key]))) as $ip) {
                    $ip = trim($ip);
                    if (filter_var($ip, FILTER_VALIDATE_IP, FILTER_FLAG_NO_PRIV_RANGE | FILTER_FLAG_NO_RES_RANGE) !== false) {
                        return $ip;
                    }
                }
            }
        }
        return isset($_SERVER['REMOTE_ADDR']) ? sanitize_text_field(wp_unslash($_SERVER['REMOTE_ADDR'])) : 'Unknown';
    }

    private function render_key_management_section() {
        global $wpdb;
        
        // Get current active key
        $keys_table = $wpdb->prefix . 'gst_audit_keys';
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        $current_key = $wpdb->get_var("SELECT key_value FROM " . $keys_table . " WHERE is_active = 1 ORDER BY created_at DESC LIMIT 1");
        
        if (!$current_key) {
            // Generate initial key if none exists
            $current_key = wp_generate_password(32, false);
            // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
            $wpdb->insert(
                $keys_table,
                [
                    'key_value' => $current_key,
                    'is_active' => 1,
                    'created_at' => current_time('mysql')
                ],
                ['%s', '%d', '%s']
            );
            update_option('gst_audit_email_secret_key', $current_key);
            
            // Clear cache after initial key creation
            wp_cache_delete('gst_audit_key_history', 'gst_audit');
        }
        
        $public_url = home_url('/?gst_trigger_email=1&key=' . $current_key);
        
        // Get key history with caching
        $cache_key = 'gst_audit_key_history';
        $key_history = wp_cache_get($cache_key, 'gst_audit');
        
        if (false === $key_history) {
            // phpcs:disable WordPress.DB.DirectDatabaseQuery.DirectQuery, WordPress.DB.DirectDatabaseQuery.NoCaching, WordPress.DB.PreparedSQL.NotPrepared
            $key_history = $wpdb->get_results(
                "SELECT k.*, 
                        (SELECT COUNT(*) FROM {$wpdb->prefix}gst_audit_key_logs l WHERE l.key_id = k.id) as access_count,
                        (SELECT MAX(l.access_time) FROM {$wpdb->prefix}gst_audit_key_logs l WHERE l.key_id = k.id) as last_access
                 FROM " . $keys_table . " k 
                 ORDER BY k.created_at DESC 
                 LIMIT 10"
            );
            // phpcs:enable WordPress.DB.DirectDatabaseQuery.DirectQuery, WordPress.DB.DirectDatabaseQuery.NoCaching, WordPress.DB.PreparedSQL.NotPrepared
            
            // Cache for 5 minutes
            wp_cache_set($cache_key, $key_history, 'gst_audit', 300);
        }
        
        echo '<div class="notice notice-info">';
        echo '<h4>Public Email Trigger URL - Key Management</h4>';
        
        // Current Key Section
        echo '<div style="background: #f9f9f9; padding: 15px; margin: 10px 0; border-radius: 4px;">';
        echo '<h5>Current Active Key</h5>';
        echo '<p><strong>URL:</strong> <code id="current-url">' . esc_html($public_url) . '</code> <button type="button" id="copy-url-btn" class="button button-small" style="margin-left:10px;">Copy URL</button></p>';
        echo '<p><strong>Key:</strong> <code id="current-key">' . esc_html($current_key) . '</code> <button type="button" id="copy-key-btn" class="button button-small" style="margin-left:10px;">Copy Key</button></p>';
        echo '<button type="button" id="regenerate-key-btn" class="button button-secondary">Regenerate Key</button>';
        echo '<span id="regenerate-spinner" style="display:none; margin-left:10px;"><span class="spinner is-active" style="float:none;display:inline-block;"></span> Regenerating...</span>';
        echo '</div>';
        
        // Key History Table
        if (!empty($key_history)) {
            echo '<h5>Key History</h5>';
            echo '<table class="wp-list-table widefat fixed striped">';
            echo '<thead>';
            echo '<tr>';
            echo '<th>Key (First 8 chars)</th>';
            echo '<th>Status</th>';
            echo '<th>Created</th>';
            echo '<th>Access Count</th>';
            echo '<th>Last Access</th>';
            echo '<th>Actions</th>';
            echo '</tr>';
            echo '</thead>';
            echo '<tbody>';
            
            foreach ($key_history as $key) {
                $key_preview = substr($key->key_value, 0, 8) . '...';
                $status = $key->is_active ? '<span style="color: green;">Active</span>' : '<span style="color: red;">Inactive</span>';
                $created = $key->created_at ? gmdate('Y-m-d H:i:s', strtotime($key->created_at)) : 'N/A';
                $last_access = $key->last_access ? gmdate('Y-m-d H:i:s', strtotime($key->last_access)) : 'Never';
                
                echo '<tr>';
                echo '<td><code>' . esc_html($key_preview) . '</code></td>';
                echo '<td>' . esc_html($status) . '</td>';
                echo '<td>' . esc_html($created) . '</td>';
                echo '<td>' . esc_html($key->access_count) . '</td>';
                echo '<td>' . esc_html($last_access) . '</td>';
                echo '<td>';
                echo '<button type="button" class="button button-small view-logs-btn" data-key-id="' . esc_attr($key->id) . '">View Logs</button>';
                echo '</td>';
                echo '</tr>';
            }
            
            echo '</tbody>';
            echo '</table>';
        }
        
        echo '<p><strong>Usage:</strong> Access the current URL to manually trigger the monthly email without authentication.</p>';
        echo '<p><strong>Security:</strong> The URL contains a secret key. Keep it private and do not share it publicly.</p>';
        echo '<p><strong>External Cron:</strong> You can use this URL with external cron services or server cron jobs.</p>';
        echo '<p><strong>Test:</strong> <a href="' . esc_url($public_url) . '" target="_blank" class="button button-secondary">Test Public URL</a></p>';
        echo '</div>';
        
        // Add JavaScript for key management
        echo '<script>
        jQuery(document).ready(function($) {
            console.log("Webhook copy functionality loaded");
            
            // Copy URL functionality
            $(document).on("click", "#copy-url-btn", function(e) {
                e.preventDefault();
                console.log("Copy URL button clicked");
                var url = $("#current-url").text();
                console.log("URL to copy:", url);
                copyToClipboard(url, this, "URL copied to clipboard!");
            });
            
            // Copy Key functionality
            $(document).on("click", "#copy-key-btn", function(e) {
                e.preventDefault();
                console.log("Copy Key button clicked");
                var key = $("#current-key").text();
                console.log("Key to copy:", key);
                copyToClipboard(key, this, "Key copied to clipboard!");
            });
            
            // Copy to clipboard function
            function copyToClipboard(text, button, successMessage) {
                console.log("Attempting to copy:", text);
                
                // Try modern clipboard API first
                if (navigator.clipboard && window.isSecureContext) {
                    console.log("Using modern clipboard API");
                    navigator.clipboard.writeText(text).then(function() {
                        console.log("Copy successful with modern API");
                        showCopySuccess(button, successMessage);
                    }).catch(function(err) {
                        console.log("Modern API failed, using fallback:", err);
                        fallbackCopyTextToClipboard(text, button, successMessage);
                    });
                } else {
                    console.log("Using fallback copy method");
                    fallbackCopyTextToClipboard(text, button, successMessage);
                }
            }
            
            // Fallback copy function for older browsers
            function fallbackCopyTextToClipboard(text, button, successMessage) {
                var textArea = document.createElement("textarea");
                textArea.value = text;
                textArea.style.position = "fixed";
                textArea.style.left = "-999999px";
                textArea.style.top = "-999999px";
                document.body.appendChild(textArea);
                textArea.focus();
                textArea.select();
                
                try {
                    var successful = document.execCommand("copy");
                    console.log("Fallback copy result:", successful);
                    if (successful) {
                        showCopySuccess(button, successMessage);
                    } else {
                        showCopyError("Failed to copy to clipboard");
                    }
                } catch (err) {
                    console.log("Fallback copy error:", err);
                    showCopyError("Failed to copy to clipboard");
                }
                
                document.body.removeChild(textArea);
            }
            
            // Show copy success message
            function showCopySuccess(button, message) {
                console.log("Showing copy success");
                var originalText = button.textContent;
                button.textContent = "Copied!";
                button.style.backgroundColor = "#46b450";
                button.style.color = "white";
                
                setTimeout(function() {
                    button.textContent = originalText;
                    button.style.backgroundColor = "";
                    button.style.color = "";
                }, 2000);
            }
            
            // Show copy error message
            function showCopyError(message) {
                console.log("Showing copy error:", message);
                alert(message);
            }
            
            $("#regenerate-key-btn").click(function() {
                if (!confirm("Are you sure you want to regenerate the key? The old key will become inactive.")) {
                    return;
                }
                
                $("#regenerate-spinner").show();
                $("#regenerate-key-btn").prop("disabled", true);
                
                $.ajax({
                    url: ajaxurl,
                    type: "POST",
                    data: {
                        action: "gst_audit_regenerate_key",
                        nonce: "' . esc_js(wp_create_nonce('gst_audit_mail_settings')) . '"
                    },
                    success: function(response) {
                        if (response.success) {
                            location.reload();
                        } else {
                            alert("Error: " + response.data);
                        }
                    },
                    error: function() {
                        alert("Error regenerating key. Please try again.");
                    },
                    complete: function() {
                        $("#regenerate-spinner").hide();
                        $("#regenerate-key-btn").prop("disabled", false);
                    }
                });
            });
            
            $(".view-logs-btn").click(function() {
                var keyId = $(this).data("key-id");
                var url = "' . esc_js(admin_url('admin.php')) . '?page=gst-audit-export&tab=webhook&view_logs=" + keyId;
                console.log("Opening URL:", url);
                window.open(url, "_blank");
            });
        });
        </script>';
    }

    private function render_key_logs($key_id) {
        global $wpdb;
        
        $keys_table = $wpdb->prefix . 'gst_audit_keys';
        $logs_table = $wpdb->prefix . 'gst_audit_key_logs';
        
        // Get key information
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.InterpolatedNotPrepared - Table name cannot be parameterized
        $key_info = $wpdb->get_row($wpdb->prepare(
            "SELECT * FROM " . $keys_table . " WHERE id = %d", // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared
            $key_id
        ));
        
        if (!$key_info) {
            echo '<div class="notice notice-error"><p>Key not found.</p></div>';
            return;
        }
        
        // Get logs for this key
        // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
        // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.InterpolatedNotPrepared - Table name cannot be parameterized
        $logs = $wpdb->get_results($wpdb->prepare(
            "SELECT * FROM " . $logs_table . " WHERE key_id = %d ORDER BY access_time DESC LIMIT 100", // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared
            $key_id
        ));
        
        echo '<div class="notice notice-info">';
        echo '<h4>Access Logs for Key: ' . esc_html(substr($key_info->key_value, 0, 8)) . '...</h4>';
        echo '<p><a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=webhook')) . '" class="button button-secondary">&larr; Back to Key Management</a></p>';
        echo '</div>';
        
        if (empty($logs)) {
            echo '<div class="notice notice-warning"><p>No access logs found for this key.</p></div>';
            return;
        }
        
        echo '<table class="wp-list-table widefat fixed striped">';
        echo '<thead>';
        echo '<tr>';
        echo '<th>Access Time</th>';
        echo '<th>IP Address</th>';
        echo '<th>Browser</th>';
        echo '<th>OS</th>';
        echo '<th>Location</th>';
        echo '<th>Status</th>';
        echo '<th>Error Message</th>';
        echo '</tr>';
        echo '</thead>';
        echo '<tbody>';
        
        foreach ($logs as $log) {
            $access_time = $log->access_time ? gmdate('Y-m-d H:i:s', strtotime($log->access_time)) : 'N/A';
            $status = $log->success ? '<span style="color: green;">Success</span>' : '<span style="color: red;">Failed</span>';
            $error_msg = $log->error_message ? $log->error_message : '-';
            
            echo '<tr>';
            echo '<td>' . esc_html($access_time) . '</td>';
            echo '<td>' . esc_html($log->ip_address) . '</td>';
            echo '<td>' . esc_html($log->browser) . '</td>';
            echo '<td>' . esc_html($log->os) . '</td>';
            echo '<td>' . esc_html($log->location) . '</td>';
            echo '<td>' . wp_kses($status, ['span' => ['style' => []]]) . '</td>';
            echo '<td>' . esc_html($error_msg) . '</td>';
            echo '</tr>';
        }
        
        echo '</tbody>';
        echo '</table>';
        
        echo '<p><em>Showing last 100 access attempts. Total access count: ' . count($logs) . '</em></p>';
    }

    private function render_webhook_settings() {
        // Debug information
        echo '<div class="notice notice-info">';
        echo '<p><strong>Debug Info:</strong></p>';
        echo '<p>Current URL: ' . esc_html(sanitize_text_field(wp_unslash($_SERVER['REQUEST_URI'] ?? ''))) . '</p>';
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for debugging are safe
        echo '<p>Page parameter: ' . esc_html($_GET['page'] ?? 'not set') . '</p>';
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for debugging are safe
        echo '<p>Tab parameter: ' . esc_html($_GET['tab'] ?? 'not set') . '</p>';
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for debugging are safe
        echo '<p>View logs parameter: ' . esc_html($_GET['view_logs'] ?? 'not set') . '</p>';
        echo '</div>';
        
        // Check if viewing logs for a specific key
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for navigation are safe
        if (isset($_GET['view_logs']) && !empty($_GET['view_logs'])) {
            // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for navigation are safe
            $this->render_key_logs($_GET['view_logs']);
        } else {
            // Key Management Section
            $this->render_key_management_section();
        }
    }

    private function get_hsn_code($product_id, $variation_id = null) {
        if (!$product_id) {
            return '';
        }
        
        // Try multiple sources for HSN code
        $hsn_code = '';
        
        // Always prioritize parent product HSN code over variations
        $check_ids = [];
        $check_ids[] = $product_id; // Check parent product first
        
        // Only check variation if parent has no HSN code
        if ($variation_id && $variation_id != $product_id) {
            $check_ids[] = $variation_id; // Check variation as fallback
        }
        
        foreach ($check_ids as $id) {
            // 1. Primary source: hsn_code meta
            $hsn_code = get_post_meta($id, 'hsn_code', true);
            if (!empty($hsn_code)) {
                return $hsn_code;
            }
            
            // 2. Alternative source: _hsn_code meta (some plugins use underscore prefix)
            $hsn_code = get_post_meta($id, '_hsn_code', true);
            if (!empty($hsn_code)) {
                return $hsn_code;
            }
            
            // 3. Check for HSN in product meta data (some themes/plugins store it differently)
            $meta_keys = ['_product_hsn_code', 'product_hsn_code', 'hsn', '_hsn'];
            foreach ($meta_keys as $key) {
                $hsn_code = get_post_meta($id, $key, true);
                if (!empty($hsn_code)) {
                    return $hsn_code;
                }
            }
            
            // 4. Check for HSN in product custom fields
            $custom_fields = get_post_custom($id);
            foreach ($custom_fields as $key => $value) {
                if (strpos(strtolower($key), 'hsn') !== false && !empty($value[0])) {
                    return $value[0];
                }
            }
        }
        
        // 5. Check for HSN in product attributes (only check parent product for attributes)
        $product = wc_get_product($product_id);
        if ($product) {
            $attributes = $product->get_attributes();
            foreach ($attributes as $attribute) {
                if (is_a($attribute, 'WC_Product_Attribute')) {
                    $name = $attribute->get_name();
                    if (strpos(strtolower($name), 'hsn') !== false) {
                        $values = $attribute->get_options();
                        if (!empty($values)) {
                            return $values[0];
                        }
                    }
                }
            }
        }
        
        return '';
    }

    public function save_mail_settings() {
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (!isset($_POST['gst_audit_mail_nonce']) || !wp_verify_nonce(wp_unslash($_POST['gst_audit_mail_nonce']), 'gst_audit_mail_settings')) {
            wp_die(esc_html(__('Security check failed. Please try again.', 'gst-audit-exporter-for-woocommerce')));
        }
        
        $mail_enabled = isset($_POST['mail_enabled']) ? 'yes' : 'no';
        $mail_recipient_raw = isset($_POST['mail_recipient']) ? sanitize_text_field(wp_unslash($_POST['mail_recipient'])) : '';
        $mail_day = isset($_POST['mail_day']) ? intval(wp_unslash($_POST['mail_day'])) : 1;
        $mail_time = isset($_POST['mail_time']) ? sanitize_text_field(wp_unslash($_POST['mail_time'])) : '09:00';
        
        // Validate and sanitize multiple email addresses
        $mail_recipient = $this->validate_and_sanitize_emails($mail_recipient_raw);
        
        // Validate day (1-28)
        if ($mail_day < 1 || $mail_day > 28) {
            $mail_day = 1;
        }
        
        // Validate time format (HH:MM)
        if (!preg_match('/^([01]?[0-9]|2[0-3]):[0-5][0-9]$/', $mail_time)) {
            $mail_time = '09:00';
        }
        
        update_option('gst_audit_mail_enabled', $mail_enabled);
        update_option('gst_audit_mail_recipient', $mail_recipient);
        update_option('gst_audit_mail_day', $mail_day);
        update_option('gst_audit_mail_time', $mail_time);
        
        // Reschedule cron if enabled
        if ($mail_enabled === 'yes' && !empty($mail_recipient)) {
            $this->schedule_monthly_email();
            $next_scheduled = wp_next_scheduled('gst_audit_monthly_email');
            if ($next_scheduled) {
                $next_date = gmdate('Y-m-d H:i:s', $next_scheduled);
                echo '<div class="notice notice-success"><p>Mail settings saved successfully! Next email scheduled for: ' . esc_html($next_date) . '</p></div>';
            } else {
                echo '<div class="notice notice-warning"><p>Mail settings saved, but scheduling failed. Please check your WordPress cron configuration.</p></div>';
            }
        } else {
            $this->unschedule_monthly_email();
            echo '<div class="notice notice-success"><p>Mail settings saved successfully! Automatic emails disabled.</p></div>';
        }
    }

    public function send_test_mail() {
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (!isset($_POST['gst_audit_mail_nonce']) || !wp_verify_nonce(wp_unslash($_POST['gst_audit_mail_nonce']), 'gst_audit_mail_settings')) {
            wp_die(esc_html(__('Security check failed. Please try again.', 'gst-audit-exporter-for-woocommerce')));
        }
        
        $mail_recipient = get_option('gst_audit_mail_recipient', '');
        
        if (empty($mail_recipient)) {
            echo '<div class="notice notice-error"><p>Please enter an email recipient first.</p></div>';
            return;
        }
        
        // Get previous month's data
        $previous_month = gmdate('Y-m', strtotime('first day of last month'));
        $generation_datetime = current_time('Y-m-d H:i:s');
        
        // Generate Excel file with generation date/time
        $excel_data = $this->generate_excel_data($previous_month, $generation_datetime);
        
        if (empty($excel_data)) {
            echo '<div class="notice notice-error"><p>No data found for the previous month (' . esc_html($previous_month) . ').</p></div>';
            return;
        }
        
        // Create temporary file
        $upload_dir = wp_upload_dir();
        $filename = 'gst-audit-test-export-' . $previous_month . '.xlsx';
        $file_path = $upload_dir['path'] . '/' . $filename;
        
        file_put_contents($file_path, $excel_data);
        
        // Send test email to multiple recipients
        $subject = 'TEST - Monthly GST Report - ' . $previous_month;
        $message = 'This is a test email. Please find attached the monthly GST report for ' . $previous_month . '.<br><br>';
        $message .= '<strong>Report generated on:</strong> ' . $generation_datetime . '<br><br>';
        $message .= '<strong>Note:</strong> This is a test email sent from the GST Audit Exporter plugin.';
        
        $headers = array('Content-Type: text/html; charset=UTF-8');
        $attachments = array($file_path);
        
        // Split recipients and send to each
        $recipients = array_map('trim', explode(',', $mail_recipient));
        $sent = true;
        
        foreach ($recipients as $recipient) {
            if (!empty($recipient)) {
                $result = wp_mail($recipient, $subject, $message, $headers, $attachments);
                if (!$result) {
                    $sent = false;
                }
            }
        }
        
        // Clean up temporary file
        if (file_exists($file_path)) {
            wp_delete_file($file_path);
        }
        
        if ($sent) {
            echo '<div class="notice notice-success"><p>Test email sent successfully to ' . esc_html($mail_recipient) . '!</p></div>';
        } else {
            echo '<div class="notice notice-error"><p>Failed to send test email. Please check your email configuration.</p></div>';
        }
    }

    public function trigger_manual_email() {
        // @codingStandardsIgnoreLine WordPress.Security.ValidatedSanitizedInput.InputNotSanitized - Nonce verification handles sanitization internally
        if (!isset($_POST['gst_audit_mail_nonce']) || !wp_verify_nonce(wp_unslash($_POST['gst_audit_mail_nonce']), 'gst_audit_mail_settings')) {
            wp_die(esc_html(__('Security check failed. Please try again.', 'gst-audit-exporter-for-woocommerce')));
        }
        
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        $mail_recipient = get_option('gst_audit_mail_recipient', '');
        
        if ($mail_enabled !== 'yes' || empty($mail_recipient)) {
            echo '<div class="notice notice-error"><p>Please enable automatic emails and enter a recipient email first.</p></div>';
            return;
        }
        
        // Trigger the monthly email function
        $this->send_monthly_gst_report();
        echo '<div class="notice notice-success"><p>Manual email triggered successfully! Check your email inbox.</p></div>';
    }


    public function ajax_save_mail_settings() {
        check_ajax_referer('gst_audit_mail_settings', 'nonce');
        
        $mail_enabled = isset($_POST['mail_enabled']) ? 'yes' : 'no';
        $mail_recipient_raw = isset($_POST['mail_recipient']) ? sanitize_text_field(wp_unslash($_POST['mail_recipient'])) : '';
        $mail_day = isset($_POST['mail_day']) ? intval(wp_unslash($_POST['mail_day'])) : 1;
        $mail_time = isset($_POST['mail_time']) ? sanitize_text_field(wp_unslash($_POST['mail_time'])) : '09:00';
        
        // Validate and sanitize multiple email addresses
        $mail_recipient = $this->validate_and_sanitize_emails($mail_recipient_raw);
        
        if ($mail_day < 1 || $mail_day > 28) {
            $mail_day = 1;
        }
        
        // Validate time format (HH:MM)
        if (!preg_match('/^([01]?[0-9]|2[0-3]):[0-5][0-9]$/', $mail_time)) {
            $mail_time = '09:00';
        }
        
        update_option('gst_audit_mail_enabled', $mail_enabled);
        update_option('gst_audit_mail_recipient', $mail_recipient);
        update_option('gst_audit_mail_day', $mail_day);
        update_option('gst_audit_mail_time', $mail_time);
        
        if ($mail_enabled === 'yes' && !empty($mail_recipient)) {
            $this->schedule_monthly_email();
        } else {
            $this->unschedule_monthly_email();
        }
        
        wp_send_json_success(['message' => 'Mail settings saved successfully!']);
    }

    public function schedule_monthly_email() {
        // Clear existing schedule
        $this->unschedule_monthly_email();
        
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        $mail_day = get_option('gst_audit_mail_day', '1');
        $mail_time = get_option('gst_audit_mail_time', '09:00');
        
        if ($mail_enabled === 'yes') {
            // Parse time
            $time_parts = explode(':', $mail_time);
            $hour = intval($time_parts[0]);
            $minute = intval($time_parts[1]);
            
            // Calculate next run time
            $current_time = current_time('timestamp');
            $next_run = mktime($hour, $minute, 0, gmdate('n', $current_time), $mail_day, gmdate('Y', $current_time));
            
            // If the day has passed this month, schedule for next month
            if ($next_run <= $current_time) {
                $next_run = mktime($hour, $minute, 0, gmdate('n', $current_time) + 1, $mail_day, gmdate('Y', $current_time));
            }
            
            // Schedule the event with multiple attempts
            $scheduled = wp_schedule_single_event($next_run, 'gst_audit_monthly_email');
            
            // If scheduling failed, try multiple fallbacks
            if ($scheduled === false) {
                // Try 5 minutes from now
                $fallback_time = time() + 120;
                $scheduled = wp_schedule_single_event($fallback_time, 'gst_audit_monthly_email');
                
                // If still failed, try 1 minute from now
                if ($scheduled === false) {
                    $fallback_time = time() + 60;
                    wp_schedule_single_event($fallback_time, 'gst_audit_monthly_email');
                }
            }
            
            // Also ensure the 5-minute check is scheduled
            $this->schedule_minute_check();
        }
    }

    public function unschedule_monthly_email() {
        wp_clear_scheduled_hook('gst_audit_monthly_email');
    }

    public function send_monthly_gst_report() {
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        $mail_recipient = get_option('gst_audit_mail_recipient', '');
        
        if ($mail_enabled !== 'yes' || empty($mail_recipient)) {
            return;
        }
        
        // Get previous month's data
        $previous_month = gmdate('Y-m', strtotime('first day of last month'));
        $generation_datetime = current_time('Y-m-d H:i:s');
        
        // Generate Excel file with generation date/time
        $excel_data = $this->generate_excel_data($previous_month, $generation_datetime);
        
        if (empty($excel_data)) {
            return;
        }
        
        // Create temporary file
        $upload_dir = wp_upload_dir();
        $filename = 'gst-audit-export-' . $previous_month . '.xlsx';
        $file_path = $upload_dir['path'] . '/' . $filename;
        
        file_put_contents($file_path, $excel_data);
        
        // Send email to multiple recipients
        $subject = 'Monthly GST Report - ' . $previous_month;
        $message = 'Please find attached the monthly GST report for ' . $previous_month . '.<br><br><strong>Report generated on:</strong> ' . $generation_datetime;
        
        $headers = array('Content-Type: text/html; charset=UTF-8');
        $attachments = array($file_path);
        
        // Split recipients and send to each
        $recipients = array_map('trim', explode(',', $mail_recipient));
        $sent = true;
        
        foreach ($recipients as $recipient) {
            if (!empty($recipient)) {
                $result = wp_mail($recipient, $subject, $message, $headers, $attachments);
                if (!$result) {
                    $sent = false;
                }
            }
        }
        
        // Clean up temporary file
        if (file_exists($file_path)) {
            wp_delete_file($file_path);
        }
        
        // Schedule next month's email
        $this->schedule_monthly_email();
    }

    private function generate_excel_data($selected_month, $generation_datetime = null) {
        require_once __DIR__ . '/vendor/autoload.php';
        if (!class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet')) {
            return false;
        }
        
        global $wpdb;
        $start_date = $selected_month . '-01 00:00:00';
        $end_date = gmdate('Y-m-t 23:59:59', strtotime($start_date));
        $args = array(
            'status' => array('wc-completed', 'wc-processing'),
            'type'   => 'shop_order',
            'date_query' => array(
                array(
                    'after'     => $start_date,
                    'before'    => $end_date,
                    'inclusive' => true,
                ),
            ),
            'limit' => -1,
        );
        $all_orders = wc_get_orders($args);
        $class_rates = $this->get_tax_classes_and_rates();
        
        // Build header rows
        $header1 = ['Order Date','Order ID','Invoice Number','Order Status','Name','City','Pincode','Product Name','HSN Code','Price (Inc Tax)','Qty','Total (Inc Tax)','Total Price (Excl Tax)'];
        $header2 = ['', '', '', '', '', '', '', '', '', '', '', '', ''];
        foreach ($class_rates as $class_name => $rates) {
            $colspan = max(1, count($rates));
            $header1[] = $class_name;
            for ($i = 1; $i < $colspan; $i++) {
                $header1[] = '';
            }
            if (empty($rates)) {
                $header2[] = '-';
            } else {
                foreach ($rates as $rate) {
                    $label = $rate['label'] ? $rate['label'] : 'Rate';
                    $percent = $rate['percent'] !== '' ? ' (' . $rate['percent'] . '%)' : '';
                    $header2[] = $label . $percent;
                }
            }
        }
        
        $rows = [];
        $rows[] = $header1;
        $rows[] = $header2;
        
        foreach ($all_orders as $order) {
            $order_date = $order->get_date_created() ? $order->get_date_created()->date('Y-m-d') : '';
            $order_id = $order->get_id();
            $invoice_number = $order->get_meta('_wcpdf_invoice_number');
            $order_status = $order->get_status();
            $order_name = $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
            $order_city = $order->get_billing_city();
            $order_pincode = $order->get_billing_postcode();
            
            // Product rows
            foreach ($order->get_items() as $item) {
                $product = $item->get_product();
                $product_name = $item->get_name();
                // For variations, get the parent product ID for HSN code
                $variation_id = $item->get_variation_id();
                $product_id = $product ? $product->get_id() : 0;
                
                // If this is a variation, get the parent product ID
                if ($variation_id && $variation_id > 0) {
                    $variation_product = wc_get_product($variation_id);
                    if ($variation_product && $variation_product->get_parent_id()) {
                        $product_id = $variation_product->get_parent_id();
                    }
                }
                
                $hsn_code = $product ? $this->get_hsn_code($product_id, $variation_id) : '';
                $price_inc_tax = wc_format_decimal($order->get_item_total($item, true, false), 2);
                $qty = $item->get_quantity();
                $total_excl_tax = wc_format_decimal($order->get_line_total($item, false, false), 2);
                $line_total_inc_tax = wc_format_decimal($price_inc_tax * $qty, 2);
                
                $tax_amounts = [];
                foreach ($class_rates as $class_name => $rates) {
                    if (empty($rates)) {
                        $tax_amounts[$class_name] = ['-'];
                    } else {
                        foreach ($rates as $rate) {
                            $tax_amounts[$class_name][$rate['id']] = '';
                        }
                    }
                }
                
                $item_taxes = $item->get_taxes();
                if (!empty($item_taxes['total'])) {
                    foreach ($item_taxes['total'] as $tax_rate_id => $tax_amount) {
                        foreach ($class_rates as $class_name => $rates) {
                            foreach ($rates as $rate) {
                                if ($rate['id'] == $tax_rate_id) {
                                    $tax_amounts[$class_name][$tax_rate_id] = wc_format_decimal($tax_amount, 2);
                                }
                            }
                        }
                    }
                }
                
                $row = [$order_date, $order_id, $invoice_number, $order_status, $order_name, $order_city, $order_pincode, $product_name, $hsn_code, $price_inc_tax, $qty, $line_total_inc_tax, $total_excl_tax];
                foreach ($class_rates as $class_name => $rates) {
                    if (empty($rates)) {
                        $row[] = '-';
                    } else {
                        foreach ($rates as $rate) {
                            $row[] = isset($tax_amounts[$class_name][$rate['id']]) ? $tax_amounts[$class_name][$rate['id']] : '';
                        }
                    }
                }
                $rows[] = $row;
            }
            
            // Shipping row
            $shipping_total = 0;
            $shipping_total_tax = 0;
            $shipping_tax_amounts = [];
            foreach ($class_rates as $class_name => $rates) {
                if (empty($rates)) {
                    $shipping_tax_amounts[$class_name] = ['-'];
                } else {
                    foreach ($rates as $rate) {
                        $shipping_tax_amounts[$class_name][$rate['id']] = '';
                    }
                }
            }
            
            foreach ($order->get_items('shipping') as $shipping_item) {
                $shipping_total += (float) $shipping_item->get_total();
                $shipping_total_tax += (float) $shipping_item->get_total_tax();
                $shipping_taxes = $shipping_item->get_taxes();
                if (!empty($shipping_taxes['total'])) {
                    foreach ($shipping_taxes['total'] as $tax_rate_id => $tax_amount) {
                        foreach ($class_rates as $class_name => $rates) {
                            foreach ($rates as $rate) {
                                if ($rate['id'] == $tax_rate_id) {
                                    $prev = isset($shipping_tax_amounts[$class_name][$tax_rate_id]) && $shipping_tax_amounts[$class_name][$tax_rate_id] !== '' ? (float)$shipping_tax_amounts[$class_name][$tax_rate_id] : 0;
                                    $shipping_tax_amounts[$class_name][$tax_rate_id] = wc_format_decimal($prev + (float)$tax_amount, 2);
                                }
                            }
                        }
                    }
                }
            }
            
            if ($shipping_total > 0 || $shipping_total_tax > 0) {
                $row = [
                    $order_date,
                    $order_id,
                    $invoice_number,
                    $order_status,
                    $order_name,
                    $order_city,
                    $order_pincode,
                    'Shipping',
                    '',
                    wc_format_decimal($shipping_total + $shipping_total_tax, 2),
                    1,
                    wc_format_decimal($shipping_total + $shipping_total_tax, 2),
                    wc_format_decimal($shipping_total, 2)
                ];
                foreach ($class_rates as $class_name => $rates) {
                    if (empty($rates)) {
                        $row[] = '-';
                    } else {
                        foreach ($rates as $rate) {
                            $row[] = isset($shipping_tax_amounts[$class_name][$rate['id']]) ? $shipping_tax_amounts[$class_name][$rate['id']] : '';
                        }
                    }
                }
                $rows[] = $row;
            }
        }
        
        // Create spreadsheet
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        // Add generation date/time header if provided
        $header_offset = 0;
        if ($generation_datetime) {
            $sheet->setCellValue('A1', 'Report Generated On: ' . $generation_datetime);
            $sheet->getStyle('A1')->getFont()->setBold(true);
            $sheet->getStyle('A1')->getFont()->setSize(12);
            $header_offset = 1;
        }
        
        // Write data
        foreach ($rows as $rowIdx => $row) {
            foreach ($row as $colIdx => $value) {
                $cell = $sheet->getCellByColumnAndRow($colIdx + 1, $rowIdx + 1 + $header_offset);
                $cell->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            }
        }
        
        // Generate Excel content
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        ob_start();
        $writer->save('php://output');
        return ob_get_clean();
    }

    public function ensure_cron_triggered() {
        // Only run on admin pages to avoid performance issues
        if (!is_admin()) {
            return;
        }
        
        // Check if we need to reschedule
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        if ($mail_enabled === 'yes') {
            $next_scheduled = wp_next_scheduled('gst_audit_monthly_email');
            $minute_scheduled = wp_next_scheduled('gst_audit_minute_check');
            
            if (!$next_scheduled) {
                // Reschedule if no event is scheduled
                $this->schedule_monthly_email();
            }
            
            if (!$minute_scheduled) {
                // Ensure 5-minute check is scheduled
                $this->schedule_minute_check();
            }
        }
    }

    public function schedule_hourly_check() {
        // Schedule hourly check if not already scheduled
        if (!wp_next_scheduled('gst_audit_hourly_check')) {
            wp_schedule_event(time(), 'hourly', 'gst_audit_hourly_check');
        }
    }

    public function schedule_minute_check() {
        // Schedule 5-minute check if not already scheduled
        if (!wp_next_scheduled('gst_audit_minute_check')) {
            wp_schedule_event(time(), 'gst_audit_every_5_minutes', 'gst_audit_minute_check');
        }
    }

    public function add_custom_cron_intervals($schedules) {
        $schedules['gst_audit_every_5_minutes'] = array(
            'interval' => 120, // 5 minutes in seconds
            'display' => __('Every 5 Minutes', 'gst-audit-exporter-for-woocommerce')
        );
        return $schedules;
    }

    public function aggressive_cron_check() {
        // Only run if mail is enabled
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        if ($mail_enabled !== 'yes') {
            return;
        }
        
        // Check if 5-minute check is scheduled, if not schedule it
        if (!wp_next_scheduled('gst_audit_minute_check')) {
            wp_schedule_event(time(), 'gst_audit_every_5_minutes', 'gst_audit_minute_check');
        }
        
        // Check if hourly check is scheduled, if not schedule it
        if (!wp_next_scheduled('gst_audit_hourly_check')) {
            wp_schedule_event(time(), 'hourly', 'gst_audit_hourly_check');
        }
        
        // Check if monthly email is scheduled, if not schedule it
        if (!wp_next_scheduled('gst_audit_monthly_email')) {
            $this->schedule_monthly_email();
        }
    }

    public function handle_public_email_trigger() {
        // Check if this is a request to trigger the email
        // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for public email trigger are safe
        if (isset($_GET['gst_trigger_email']) && $_GET['gst_trigger_email'] === '1') {
            // @codingStandardsIgnoreLine WordPress.Security.NonceVerification.Recommended - GET parameters for public email trigger are safe
            $provided_key = $_GET['key'] ?? '';
            
            // Check if key is valid and active
            global $wpdb;
            $keys_table = $wpdb->prefix . 'gst_audit_keys';
            
            // @codingStandardsIgnoreLine WordPress.DB.DirectDatabaseQuery.DirectQuery - Custom table operation
            // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
            // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared - Table name cannot be parameterized
            // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.InterpolatedNotPrepared - Table name cannot be parameterized
            $key_data = $wpdb->get_row($wpdb->prepare(
                "SELECT * FROM " . $keys_table . " WHERE key_value = %s AND is_active = 1", // @codingStandardsIgnoreLine WordPress.DB.PreparedSQL.NotPrepared
                $provided_key
            ));
            
            if ($key_data) {
                // Valid key - send email and log success
                $this->send_monthly_gst_report();
                $this->log_key_access($provided_key, true);
                
                // Return JSON response
                header('Content-Type: application/json');
                echo json_encode([
                    'success' => true,
                    'message' => 'Monthly GST report email sent successfully',
                    'timestamp' => current_time('Y-m-d H:i:s')
                ]);
                exit;
            } else {
                // Invalid or inactive key - log failure
                $this->log_key_access($provided_key, false, 'Invalid or inactive key');
                
                // Return error for invalid key
                header('Content-Type: application/json');
                http_response_code(401);
                echo json_encode([
                    'success' => false,
                    'message' => 'Invalid or inactive secret key',
                    'timestamp' => current_time('Y-m-d H:i:s')
                ]);
                exit;
            }
        }
    }

    public function check_and_send_monthly_email() {
        $mail_enabled = get_option('gst_audit_mail_enabled', 'no');
        $mail_recipient = get_option('gst_audit_mail_recipient', '');
        $mail_day = get_option('gst_audit_mail_day', '1');
        $mail_time = get_option('gst_audit_mail_time', '09:00');
        
        if ($mail_enabled !== 'yes' || empty($mail_recipient)) {
            return;
        }
        
        // Check if it's time to send the monthly email
        $current_time = current_time('timestamp');
        $current_day = gmdate('j', $current_time);
        $current_hour = gmdate('H', $current_time);
        $current_minute = gmdate('i', $current_time);
        
        // Parse scheduled time
        $time_parts = explode(':', $mail_time);
        $scheduled_hour = intval($time_parts[0]);
        $scheduled_minute = intval($time_parts[1]);
        
        // Check if today is the scheduled day and we're at or past the scheduled time
        if ($current_day == $mail_day && 
            ($current_hour > $scheduled_hour || 
             ($current_hour == $scheduled_hour && $current_minute >= $scheduled_minute))) {
            
            // Check if we've already sent this month's email
            $last_sent_month = get_option('gst_audit_last_sent_month', '');
            $current_month = gmdate('Y-m', $current_time);
            
            if ($last_sent_month !== $current_month) {
                // Send the email
                $this->send_monthly_gst_report();
                
                // Update the last sent month
                update_option('gst_audit_last_sent_month', $current_month);
                
                // Schedule next month's email
                $this->schedule_monthly_email();
            }
        }
    }
}

new GST_Audit_Exporter(); 