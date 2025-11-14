<?php
if ( ! defined( 'ABSPATH' ) ) exit; // Exit if accessed directly
?>
<!-- GST Audit Report HSN PDF Template -->

<?php do_action( 'wpo_wcpdf_before_document', $this->get_type(), $this->order ); ?>

<table class="head container">
	<tr>
		<td class="header">
		<?php
			if ( $this->has_header_logo() ) {
				do_action( 'wpo_wcpdf_before_shop_logo', $this->get_type(), $this->order );
				$this->header_logo();
				do_action( 'wpo_wcpdf_after_shop_logo', $this->get_type(), $this->order );
			} else {
				$this->title();
			}
		?>
		</td>
		<td class="shop-info">
			<?php do_action( 'wpo_wcpdf_before_shop_name', $this->get_type(), $this->order ); ?>
			<div class="shop-name"><h3><?php $this->shop_name(); ?></h3></div>
			<?php do_action( 'wpo_wcpdf_after_shop_name', $this->get_type(), $this->order ); ?>
			<?php do_action( 'wpo_wcpdf_before_shop_address', $this->get_type(), $this->order ); ?>
			<div class="shop-address"><?php $this->shop_address(); ?></div>
			<?php do_action( 'wpo_wcpdf_after_shop_address', $this->get_type(), $this->order ); ?>
		</td>
	</tr>
</table>

<?php do_action( 'wpo_wcpdf_before_document_label', $this->get_type(), $this->order ); ?>

<?php if ( $this->has_header_logo() ) : ?>
	<h1 class="document-type-label"><?php $this->title(); ?></h1>
<?php endif; ?>

<?php do_action( 'wpo_wcpdf_after_document_label', $this->get_type(), $this->order ); ?>

<table class="order-data-addresses">
	<tr>
		<td class="address billing-address">
			<!-- <h3><?php _e( 'Billing Address:', 'woocommerce-pdf-invoices-packing-slips' ); ?></h3> -->
			<?php do_action( 'wpo_wcpdf_before_billing_address', $this->get_type(), $this->order ); ?>
			<p><?php $this->billing_address(); ?></p>
			<?php do_action( 'wpo_wcpdf_after_billing_address', $this->get_type(), $this->order ); ?>
			<?php if ( isset( $this->settings['display_email'] ) ) : ?>
				<div class="billing-email"><?php $this->billing_email(); ?></div>
			<?php endif; ?>
			<?php if ( isset( $this->settings['display_phone'] ) ) : ?>
				<div class="billing-phone"><?php $this->billing_phone(); ?></div>
			<?php endif; ?>
		</td>
		<td class="address shipping-address">
			<?php if ( $this->show_shipping_address() ) : ?>
				<h3><?php _e( 'Ship To:', 'woocommerce-pdf-invoices-packing-slips' ); ?></h3>
				<?php do_action( 'wpo_wcpdf_before_shipping_address', $this->get_type(), $this->order ); ?>
				<p><?php $this->shipping_address(); ?></p>
				<?php do_action( 'wpo_wcpdf_after_shipping_address', $this->get_type(), $this->order ); ?>
				<?php if ( isset( $this->settings['display_phone'] ) ) : ?>
					<div class="shipping-phone"><?php $this->shipping_phone(); ?></div>
				<?php endif; ?>
			<?php endif; ?>
		</td>
		<td class="order-data">
			<table>
				<?php do_action( 'wpo_wcpdf_before_order_data', $this->get_type(), $this->order ); ?>
				<?php if ( isset( $this->settings['display_number'] ) ) : ?>
					<tr class="invoice-number">
						<th><?php $this->number_title(); ?></th>
						<td><?php $this->number( $this->get_type() ); ?></td>
					</tr>
				<?php endif; ?>
				<?php if ( isset( $this->settings['display_date'] ) ) : ?>
					<tr class="invoice-date">
						<th><?php $this->date_title(); ?></th>
						<td><?php $this->date( $this->get_type() ); ?></td>
					</tr>
				<?php endif; ?>
				<tr class="order-number">
					<th><?php _e( 'Order Number:', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
					<td><?php $this->order_number(); ?></td>
				</tr>
				<tr class="order-date">
					<th><?php _e( 'Order Date:', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
					<td><?php $this->order_date(); ?></td>
				</tr>
				<?php if ( $this->get_payment_method() ) : ?>
					<tr class="payment-method">
						<th><?php _e( 'Payment Method:', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
						<td><?php $this->payment_method(); ?></td>
					</tr>
				<?php endif; ?>
				<?php do_action( 'wpo_wcpdf_after_order_data', $this->get_type(), $this->order ); ?>
			</table>
		</td>
	</tr>
</table>

<?php do_action( 'wpo_wcpdf_before_order_details', $this->get_type(), $this->order ); ?>

<table class="order-details">
	<thead>
		<tr>
			<th class="product"><?php _e( 'Product', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
			<th class="hsn_code"><?php _e( 'HSN Code', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
			<th class="quantity"><?php _e( 'Quantity', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
			<th class="price"><?php _e( 'Price', 'woocommerce-pdf-invoices-packing-slips' ); ?></th>
		</tr>
	</thead>
	<tbody>
		<?php foreach ( $this->get_order_items() as $item_id => $item ) : ?>
			<tr class="<?php echo apply_filters( 'wpo_wcpdf_item_row_class', 'item-'.$item_id, esc_attr( $this->get_type() ), $this->order, $item_id ); ?>">
				<td class="product">
					<span class="item-name"><?php echo $item['name']; ?></span>
					
					<?php do_action( 'wpo_wcpdf_before_item_meta', $this->get_type(), $item, $this->order ); ?>
					<span class="item-meta"><?php echo $item['meta']; ?></span>
					<dl class="meta">
						<?php if ( ! empty( $item['sku'] ) ) : ?><dt class="sku"><?php _e( 'SKU:', 'woocommerce-pdf-invoices-packing-slips' ); ?></dt><dd class="sku"><?php echo esc_attr( $item['sku'] ); ?></dd><?php endif; ?>
						<?php if ( ! empty( $item['weight'] ) ) : ?><dt class="weight"><?php _e( 'Weight:', 'woocommerce-pdf-invoices-packing-slips' ); ?></dt><dd class="weight"><?php echo esc_attr( $item['weight'] ); ?><?php echo esc_attr( get_option( 'woocommerce_weight_unit' ) ); ?></dd><?php endif; ?>
					</dl>
					<?php do_action( 'wpo_wcpdf_after_item_meta', $this->get_type(), $item, $this->order ); ?>
				</td>
				<td class="hsn_code">
				<?php
					$hsn_code = '';
					
					// Get product object from item array (already available from get_order_items)
					$product = isset( $item['product'] ) ? $item['product'] : null;
					
					// Fallback: get product from order item object if product not in array
					if ( ! $product && isset( $item['item'] ) && is_object( $item['item'] ) && method_exists( $item['item'], 'get_product' ) ) {
						$product = $item['item']->get_product();
					}
					
					// Fallback: get product by ID if still not available
					if ( ! $product ) {
						if ( ! empty( $item['variation_id'] ) ) {
							$product = wc_get_product( $item['variation_id'] );
						} elseif ( ! empty( $item['product_id'] ) ) {
							$product = wc_get_product( $item['product_id'] );
						}
					}
					
					if ( $product ) {
						$product_id = $product->get_id();
						
						// Get HSN code from the product using the correct meta key
						$hsn_code = get_post_meta( $product_id, 'hsn_prod_id', true );
						
						// If it's a variation and no HSN found, check parent product
						if ( empty( $hsn_code ) && $product->is_type( 'variation' ) ) {
							$parent_id = $product->get_parent_id();
							if ( $parent_id ) {
								$hsn_code = get_post_meta( $parent_id, 'hsn_prod_id', true );
							}
						}
						
						// Debug: If still no HSN, check alternative meta keys
						if ( empty( $hsn_code ) ) {
							// Try alternative meta key 'hsn_code' (used by some plugins)
							$hsn_code = get_post_meta( $product_id, 'hsn_code', true );
							
							// If still empty and it's a variation, check parent with alternative key
							if ( empty( $hsn_code ) && $product->is_type( 'variation' ) ) {
								$parent_id = $product->get_parent_id();
								if ( $parent_id ) {
									$hsn_code = get_post_meta( $parent_id, 'hsn_code', true );
								}
							}
						}
					}
					
					// Output HSN code
					if ( ! empty( $hsn_code ) ) {
						echo esc_html( $hsn_code );
					}
					?>
				</td>
				<td class="quantity"><?php echo $item['quantity']; ?></td>
				<td class="price"><?php echo $item['order_price']; ?></td>
			</tr>
		<?php endforeach; ?>
	</tbody>
	<tfoot>
		<tr class="no-borders">
					<td class="no-borders"></td>
			<td class="no-borders">
				<?php do_action( 'wpo_wcpdf_before_document_notes', $this->get_type(), $this->order ); ?>
				<div class="document-notes">
					<?php if ( $this->get_document_notes() ) : ?>
						<h3><?php _e( 'Notes', 'woocommerce-pdf-invoices-packing-slips' ); ?></h3>
						<?php $this->document_notes(); ?>
					<?php endif; ?>
				</div>
				<?php do_action( 'wpo_wcpdf_after_document_notes', $this->get_type(), $this->order ); ?>
				<?php do_action( 'wpo_wcpdf_before_customer_notes', $this->get_type(), $this->order ); ?>
				<div class="customer-notes">
					<?php if ( $this->get_shipping_notes() ) : ?>
						<h3><?php _e( 'Customer Notes', 'woocommerce-pdf-invoices-packing-slips' ); ?></h3>
						<?php $this->shipping_notes(); ?>
					<?php endif; ?>
				</div>
				<?php do_action( 'wpo_wcpdf_after_customer_notes', $this->get_type(), $this->order ); ?>
			</td>
			<td class="no-borders" colspan="2">
				<table class="totals">
					<tfoot>
						<?php foreach ( $this->get_woocommerce_totals() as $key => $total ) : ?>
							<tr class="<?php echo esc_attr( $key ); ?>">
								<th class="description"><?php echo $total['label']; ?></th>
								<td class="price"><span class="totals-price"><?php echo $total['value']; ?></span></td>
							</tr>
						<?php endforeach; ?>
					</tfoot>
				</table>
			</td>
		</tr>
	</tfoot>
</table>

<div class="bottom-spacer"></div>

<?php do_action( 'wpo_wcpdf_after_order_details', $this->get_type(), $this->order ); ?>

<?php if ( $this->get_footer() ) : ?>
	<htmlpagefooter name="docFooter"><!-- required for mPDF engine -->
		<div id="footer">
			<!-- hook available: wpo_wcpdf_before_footer -->
			<?php $this->footer(); ?>
			<!-- hook available: wpo_wcpdf_after_footer -->
		</div>
	</htmlpagefooter><!-- required for mPDF engine -->
<?php endif; ?>

<?php do_action( 'wpo_wcpdf_after_document', $this->get_type(), $this->order ); ?>

