jQuery(document).ready(function($) {
    $('#gst-audit-export-download-btn').on('click', function(e) {
        e.preventDefault();
        var $btn = $(this);
        var $spinner = $('#gst-audit-export-spinner');
        $spinner.show();
        $btn.prop('disabled', true);
        var month = $('#gst_audit_month').val();
        var perPage = $('#gst_audit_per_page').val();
        var data = new FormData();
        data.append('action', 'gst_audit_export_excel');
        data.append('gst_audit_month', month);
        data.append('gst_audit_per_page', perPage);
        data.append('_ajax_nonce', gst_audit_export.nonce);
        console.log('GST Audit Export fetch request:', Object.fromEntries(data.entries()));
        fetch(gst_audit_export.ajax_url, {
            method: 'POST',
            body: data,
            credentials: 'same-origin'
        })
        .then(function(response) {
            var contentType = response.headers.get('Content-Type') || '';
            if (response.ok && contentType.indexOf('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') !== -1) {
                return response.blob().then(function(blob) {
                    var disposition = response.headers.get('Content-Disposition');
                    var filename = 'gst-audit-export.xlsx';
                    if (disposition && disposition.indexOf('attachment') !== -1) {
                        var matches = /filename="?([^";]+)"?/.exec(disposition);
                        if (matches != null && matches[1]) filename = matches[1];
                    }
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    setTimeout(function() {
                        document.body.removeChild(a);
                        window.URL.revokeObjectURL(url);
                    }, 100);
                    $spinner.hide();
                    $btn.prop('disabled', false);
                });
            } else {
                // Try to parse error JSON
                return response.text().then(function(text) {
                    try {
                        var json = JSON.parse(text);
                        alert(json.data && json.data.message ? json.data.message : 'Failed to download file.');
                    } catch (e) {
                        alert('Failed to download file.');
                    }
                    $spinner.hide();
                    $btn.prop('disabled', false);
                });
            }
        })
        .catch(function(error) {
            console.error('GST Audit Export fetch error:', error);
            alert('Failed to download file.');
            $spinner.hide();
            $btn.prop('disabled', false);
        });
    });

    // HSN Checker AJAX save
    $(document).on('focusout', '.gst-hsn-input', function() {
        var $input = $(this);
        var productId = $input.data('product-id');
        var hsnCode = $input.val();
        var $status = $('#gst-hsn-status-' + productId);
        $status.text('Saving...').css('color', '#0073aa');
        $.ajax({
            url: gst_audit_export.ajax_url,
            method: 'POST',
            data: {
                action: 'gst_audit_save_hsn_code',
                product_id: productId,
                hsn_code: hsnCode,
                nonce: gst_audit_export.hsn_nonce
            },
            success: function(resp) {
                if (resp.success) {
                    $status.text('Saved').css('color', 'green');
                } else {
                    $status.text('Error').css('color', 'red');
                }
                setTimeout(function() { $status.text(''); }, 2000);
            },
            error: function() {
                $status.text('Error').css('color', 'red');
                setTimeout(function() { $status.text(''); }, 2000);
            }
        });
    });

    // HSN Checker AJAX pagination and per-page
    function loadHsnChecker(page, perPage) {
        $('#gst-hsn-loader').show();
        $('#gst-hsn-table-container').css('opacity', 0.5);
        $.ajax({
            url: gst_audit_export.ajax_url,
            method: 'POST',
            data: {
                action: 'gst_audit_hsn_checker',
                nonce: gst_audit_export.nonce,
                page: page,
                per_page: perPage
            },
            success: function(resp) {
                if (resp.success && resp.data && resp.data.html) {
                    $('#gst-hsn-table-container').html(resp.data.html);
                }
                $('#gst-hsn-loader').hide();
                $('#gst-hsn-table-container').css('opacity', 1);
            },
            error: function() {
                $('#gst-hsn-loader').hide();
                $('#gst-hsn-table-container').css('opacity', 1);
                alert('Failed to load HSN Checker.');
            }
        });
    }
    $(document).on('click', '.gst-hsn-page-link', function(e) {
        e.preventDefault();
        var page = $(this).data('page');
        var perPage = $('.gst-hsn-per-page-select').first().val();
        loadHsnChecker(page, perPage);
    });
    $(document).on('change', '.gst-hsn-per-page-select', function() {
        var perPage = $(this).val();
        loadHsnChecker(1, perPage);
    });
}); 