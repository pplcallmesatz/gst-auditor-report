# GST Audit Exporter for WooCommerce (v1.3.1)

A comprehensive WooCommerce extension built to simplify **GST compliance and reporting** for Indian e-commerce businesses. This powerful plugin provides automated GST reporting, HSN code management, and seamless integration with accounting workflows.

✅ **Automated Monthly GST Reports** - Email delivery with customizable scheduling  
✅ **Advanced Excel Export** - Professional GSTR-ready reports with detailed tax breakdowns  
✅ **Comprehensive HSN Management** - Bulk HSN code assignment and validation  
✅ **API & Webhook Integration** - Secure key-based automation for external systems  
✅ **Multi-source HSN Detection** - Intelligent HSN code retrieval from various plugin formats  

---

## 🛠 Developed By

**Satheesh Kumar S**  
[Github](https://github.com/pplcallmesatz/) | [Buy Me a Coffee](https://www.buymeacoffee.com/satheeshdesign)

---

## 🚀 Core Features

### 📊 **Advanced GST Export System**
- **Monthly Excel Reports**: Generate comprehensive GST reports with detailed order data
- **Multi-tax Class Support**: Automatic detection and breakdown of CGST, SGST, IGST
- **Dynamic Tax Headers**: Excel exports include all configured tax classes and rates
- **Order Status Filtering**: Export only completed/processing orders for accurate reporting
- **Date Range Selection**: Flexible monthly export with precise date filtering

### 🧾 **Intelligent HSN Code Management**
- **Bulk HSN Assignment**: Manage HSN codes for all products from a centralized interface
- **Multi-source Detection**: Automatically detects HSN codes from various plugin formats:
  - `hsn_code` and `_hsn_code` meta fields
  - Custom field variations (`_product_hsn_code`, `product_hsn_code`, `hsn`, `_hsn`)
  - Product attributes containing HSN information
  - Custom field names containing "hsn" keyword
- **Real-time AJAX Updates**: Save HSN codes without page refresh
- **Product Integration**: Direct links to edit products from the HSN checker interface

### 📧 **Automated Email System**
- **Monthly Auto-delivery**: Scheduled email reports with customizable timing
- **Multi-recipient Support**: Send reports to multiple email addresses
- **Flexible Scheduling**: Choose day of month and time for report delivery
- **Manual Triggers**: Test email functionality and send reports on-demand
- **Smart Duplicate Prevention**: Avoid sending duplicate reports for the same month

### 🔐 **Secure API & Webhook Integration**
- **Secret Key Management**: Generate and manage secure API keys for external access
- **Public URL Endpoints**: Trigger reports via secure public URLs
- **Access Logging**: Comprehensive logging of all API access attempts
- **Key History Tracking**: Monitor key usage and access patterns
- **Security Controls**: Enable/disable keys and track unauthorized access attempts

### 🎯 **Advanced Features**
- **Cron Job Management**: Reliable scheduling with multiple fallback mechanisms
- **Performance Optimization**: Efficient database queries and caching
- **Error Handling**: Comprehensive error handling with user-friendly messages
- **WordPress Standards**: Full compliance with WordPress coding standards
- **Mobile Responsive**: Works seamlessly on all devices

---

## 🚀 Installation

1. **Download** the plugin ZIP file
2. **Upload** via WordPress Admin → Plugins → Add New → Upload Plugin
3. **Activate** the plugin
4. **Navigate** to WooCommerce → GST Audit Export in the admin sidebar
5. **Configure** your settings in the 4 main tabs:
   - **Report**: Generate and export GST reports
   - **HSN Checker**: Manage HSN codes for all products
   - **Mail Settings**: Configure automated email delivery
   - **Webhook**: Set up API access and manage security keys

---

## 📤 Usage Guide

### **Generating GST Reports**
1. Go to **Report** tab
2. Select **month** for export
3. Choose **pagination** settings (optional)
4. Click **Export to Excel**
5. Download will start automatically

### **Managing HSN Codes**
1. Go to **HSN Checker** tab
2. Browse products with pagination
3. **Edit HSN codes** directly in the table
4. **Auto-save** functionality saves changes immediately
5. Click product names to **edit in WooCommerce**

### **Setting Up Automated Emails**
1. Go to **Mail Settings** tab
2. **Enable** automatic email delivery
3. Add **recipient email addresses** (comma-separated)
4. Choose **day of month** and **time** for delivery
5. **Test** the system with manual email trigger

### **API Integration**
1. Go to **Webhook** tab
2. **Generate** or **regenerate** secret keys
3. Use the **public URL** format: `yoursite.com/?gst_trigger_email=1&key=YOUR_SECRET_KEY`
4. **Monitor** access logs for security
5. **Disable** keys if needed

---

## 📋 Export Data Includes

**Order Information:**
- Order Date & ID
- Invoice Number
- Order Status
- Customer Details (Name, City, Pincode)

**Product Details:**
- Product Name
- HSN Code (auto-detected)
- Price (Including Tax)
- Quantity
- Total Amount (Including/Excluding Tax)

**Tax Breakdown:**
- All configured tax classes
- CGST, SGST, IGST rates and amounts
- Shipping tax details
- Total tax calculations

---

## 📎 Technical Requirements

- **WordPress**: 5.0+
- **WooCommerce**: 4.0+
- **PHP**: 7.2+
- **Dependencies**: PhpSpreadsheet library (included via Composer)

---

## 📌 Version History

### v1.3.1 – Current
- **Enhanced HSN Detection**: Multi-source HSN code detection from various plugin formats
- **Improved API Security**: Advanced key management with access logging
- **Better Error Handling**: Comprehensive error messages and validation
- **Performance Optimization**: Efficient database queries and caching
- **UI/UX Improvements**: Better responsive design and user experience

### v1.1.0
- Added HSN Code management section
- Improved export filters (status, payment, state)
- Performance optimization for large order sets
- Minor UI improvements

---

## 🔧 Technical Architecture

### **Database Integration**
- **Custom Tables**: Secure key management and access logging tables
- **WordPress Options**: Configuration settings and mail preferences
- **WooCommerce Integration**: Seamless order and product data access
- **Caching**: Performance optimization with WordPress object caching

### **Security Features**
- **Nonce Verification**: All forms and AJAX requests protected
- **Capability Checks**: Admin-only access with proper permission validation
- **SQL Injection Prevention**: Prepared statements for all database queries
- **XSS Protection**: Proper data sanitization and escaping
- **Secret Key Management**: Secure API key generation and validation

### **Performance Optimizations**
- **Efficient Queries**: Optimized database queries for large order sets
- **Pagination**: Memory-efficient handling of large product catalogs
- **Caching**: Strategic use of WordPress caching for key management
- **AJAX Loading**: Non-blocking UI updates for better user experience

---

## 🧑‍💻 Support & Feedback

**Developer**: Satheesh Kumar S  
📩 **Email**: `satheeshdesign@gmail.com`  
☕ **[Buy Me a Coffee](https://www.buymeacoffee.com/satheeshdesign)**  
🐛 **Issues**: [GitHub Issues](https://github.com/pplcallmesatz/gst-auditor-report/issues)  

---

## ☕ Support Development

If you find this plugin useful for your business, consider supporting its development:

[![Buy Me a Coffee](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/satheeshdesign)

Your support helps maintain and improve this plugin for the Indian e-commerce community.

---

## 📜 License

This plugin is released under the **[GPLv2 or later License](https://www.gnu.org/licenses/gpl-2.0.html)**.

You are free to:
- ✅ Use the plugin for personal and commercial projects
- ✅ Modify and customize the code
- ✅ Distribute the plugin
- ✅ Use in client projects

---

## 🤝 Contributing

Contributions are welcome! Please feel free to:
- Report bugs and issues
- Suggest new features
- Submit pull requests
- Improve documentation

---

### 🇮🇳 Built with care in India for the Indian e-commerce community

*This plugin is specifically designed to meet the GST compliance requirements of Indian businesses, making tax reporting seamless and automated.*
