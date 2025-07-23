<?php
/*
Plugin Name: GST Audit Exporter
Description: GST Audit Exporter v1.1.0 – A comprehensive WooCommerce extension to simplify GST compliance and reporting for Indian e-commerce businesses. Developed by Satheesh Kumar S at Mallow Technologies Private Limited. Export detailed order data for GST filing, manage HSN codes, and generate Excel reports.
Version: 1.1.0
Requires at least: 5.0
Requires PHP: 7.2
Author: Satheesh Kumar S 
Author URI: https://github.com/pplcallmesatz/
Website: https://github.com/pplcallmesatz/

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
    }

    public function enqueue_admin_assets($hook) {
        if ($hook !== 'toplevel_page_gst-audit-export') return;
        wp_enqueue_script('gst-audit-exporter-js', plugins_url('gst-audit-exporter.js', __FILE__), ['jquery'], null, true);
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
            wp_die( __( 'You do not have sufficient permissions to access this page.' ) );
        }
        $tab = isset($_GET['tab']) ? sanitize_text_field($_GET['tab']) : 'report';
        echo '<div class="wrap"><h1>GST Audit Export</h1>';
        echo '<h2 class="nav-tab-wrapper">';
        echo '<a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=report')) . '" class="nav-tab' . ($tab === 'report' ? ' nav-tab-active' : '') . '">Report</a>';
        echo '<a href="' . esc_url(admin_url('admin.php?page=gst-audit-export&tab=hsn')) . '" class="nav-tab' . ($tab === 'hsn' ? ' nav-tab-active' : '') . '">HSN Checker</a>';
        echo '</h2>';
        if ($tab === 'report') {
            $selected_month = isset($_POST['gst_audit_month']) ? sanitize_text_field($_POST['gst_audit_month']) : (isset($_GET['gst_audit_month']) ? sanitize_text_field($_GET['gst_audit_month']) : date('Y-m'));
            $per_page = isset($_POST['gst_audit_per_page']) ? max(10, intval($_POST['gst_audit_per_page'])) : (isset($_GET['gst_audit_per_page']) ? max(10, intval($_GET['gst_audit_per_page'])) : 20);
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
        }
        echo '</div>';
    }

    public function render_hsn_checker($paged = null, $per_page = null) {
        if ($paged === null) $paged = isset($_GET['product_page']) ? max(1, intval($_GET['product_page'])) : 1;
        if ($per_page === null) $per_page = isset($_GET['hsn_per_page']) ? max(10, intval($_GET['hsn_per_page'])) : 20;
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
                $hsn_code = get_post_meta($product_id, 'hsn_code', true);
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
        echo '<label for="' . $select_id . '">Show </label>';
        echo '<select name="hsn_per_page" id="' . $select_id . '" class="gst-hsn-per-page-select">';
        $options = [10, 20, 50, 100, 200, 500, 750, 1000];
        foreach ($options as $opt) {
            $sel = ($per_page == $opt) ? 'selected' : '';
            echo '<option value="' . $opt . '" ' . $sel . '>' . $opt . '</option>';
        }
        echo '</select> items per page';
        echo '</form>';
        echo '<strong>Total items: ' . esc_html($total_items) . '</strong> | ';
        echo 'Page ' . esc_html($paged) . ' of ' . esc_html($total_pages) . ' ';
        if ($total_pages > 1) {
            for ($i = 1; $i <= $total_pages; $i++) {
                $active = $i === $paged ? 'font-weight:bold;text-decoration:underline;' : '';
                echo '<a href="#" class="gst-hsn-page-link" data-page="' . $i . '" style="display:inline-block; margin:0 8px; padding:4px 10px; border-radius:4px; background:' . ($i === $paged ? '#dbeafe' : 'transparent') . '; color:' . ($i === $paged ? '#1d4ed8' : '#222') . '; ' . $active . '">' . $i . '</a>';
            }
        }
        echo '</div>';
    }

    public function ajax_hsn_checker() {
        check_ajax_referer('gst_audit_export_excel', 'nonce');
        $paged = isset($_POST['page']) ? max(1, intval($_POST['page'])) : 1;
        $per_page = isset($_POST['per_page']) ? max(10, intval($_POST['per_page'])) : 20;
        ob_start();
        $this->render_hsn_checker($paged, $per_page);
        $html = ob_get_clean();
        wp_send_json_success(['html' => $html]);
    }

    public function ajax_save_hsn_code() {
        check_ajax_referer('gst_audit_save_hsn_code', 'nonce');
        $product_id = isset($_POST['product_id']) ? intval($_POST['product_id']) : 0;
        $hsn_code = isset($_POST['hsn_code']) ? sanitize_text_field($_POST['hsn_code']) : '';
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
        global $wpdb;
        $tax_classes = array('Standard' => '');
        $wc_tax_classes = WC_Tax::get_tax_classes();
        foreach ($wc_tax_classes as $class) {
            $tax_classes[$class] = sanitize_title($class);
        }
        // For each class, get all rates
        $class_rates = [];
        foreach ($tax_classes as $class_name => $slug) {
            $rates = $wpdb->get_results($wpdb->prepare("SELECT * FROM {$wpdb->prefix}woocommerce_tax_rates WHERE tax_rate_class = %s ORDER BY tax_rate_order, tax_rate_priority", $slug));
            $class_rates[$class_name] = [];
            foreach ($rates as $rate) {
                $label = $rate->tax_rate_name;
                $percent = $rate->tax_rate;
                $id = $rate->tax_rate_id;
                $class_rates[$class_name][] = [
                    'id' => $id,
                    'label' => $label,
                    'percent' => $percent
                ];
            }
        }
        return $class_rates;
    }

    private function render_pagination_controls($current_page, $total_pages, $total_items, $per_page, $selected_month) {
        $base_url = remove_query_arg(['gst_audit_page', 'gst_audit_per_page']);
        echo '<div style="text-align:center; margin: 10px 0;">';
        echo '<form method="post" style="display:inline-block; margin-right:20px;">';
        echo '<input type="hidden" name="gst_audit_month" value="' . esc_attr($selected_month) . '" />';
        echo '<label for="gst_audit_per_page">Show </label>';
        echo '<select name="gst_audit_per_page" id="gst_audit_per_page" onchange="this.form.submit()">';
        $options = [10, 20, 50, 100, 200, 500, 750, 1000];
        foreach ($options as $opt) {
            $sel = ($per_page == $opt) ? 'selected' : '';
            echo '<option value="' . $opt . '" ' . $sel . '>' . $opt . '</option>';
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
        $per_page = isset($_POST['gst_audit_per_page']) ? max(10, intval($_POST['gst_audit_per_page'])) : (isset($_GET['gst_audit_per_page']) ? max(10, intval($_GET['gst_audit_per_page'])) : $default_per_page);
        $current_page = isset($_GET['gst_audit_page']) ? max(1, intval($_GET['gst_audit_page'])) : 1;
        $start_date = $selected_month . '-01 00:00:00';
        $end_date = date('Y-m-t 23:59:59', strtotime($start_date));
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
            echo '<th colspan="' . $colspan . '">' . esc_html($class_name) . '</th>';
        }
        echo '</tr>';
        // Second row: tax rate columns
        echo '<tr>';
        foreach ($class_rates as $rates) {
            if (empty($rates)) {
                echo '<th>-</th>';
            } else {
                foreach ($rates as $rate) {
                    $label = $rate['label'] ? esc_html($rate['label']) : 'Rate';
                    $percent = $rate['percent'] !== '' ? ' (' . esc_html($rate['percent']) . '%)' : '';
                    echo '<th>' . $label . $percent . '</th>';
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
            echo '<tr><td colspan="' . (9 + $total_tax_cols) . '" style="text-align:center;">No orders found for this month.</td></tr>';
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
                $hsn_code = $product ? get_post_meta($product->get_id(), 'hsn_code', true) : '';
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
        $selected_month = isset($_POST['gst_audit_month']) ? sanitize_text_field($_POST['gst_audit_month']) : date('Y-m');
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
            $end_date = date('Y-m-t 23:59:59', strtotime($start_date));
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
                    $hsn_code = $product ? get_post_meta($product->get_id(), 'hsn_code', true) : '';
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
            'label' => __('HSN Code', 'gst-audit-exporter'),
            'desc_tip' => true,
            'description' => __('Enter the HSN code for GST reporting.', 'gst-audit-exporter'),
            'type' => 'text',
        ]);
    }

    public function save_hsn_code_field($post_id) {
        $hsn_code = isset($_POST['hsn_code']) ? sanitize_text_field($_POST['hsn_code']) : '';
        update_post_meta($post_id, 'hsn_code', $hsn_code);
    }
}

new GST_Audit_Exporter(); 