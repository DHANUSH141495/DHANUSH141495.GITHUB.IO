document.addEventListener('DOMContentLoaded', function() {
    const excelBtn = document.getElementById('excelBtn');
    const toast = document.getElementById('toast');
    const toastMessage = toast.querySelector('.toast-message');

    // Configuration - customize these paths as needed
    const config = {
        // Local Excel file path (for desktop/local use)
        localFilePath: 'your-spreadsheet.xlsx',
        
        // Online Excel URL (SharePoint, OneDrive, etc.)
        onlineExcelUrl: '[office.com](https://www.office.com/launch/excel)',
        
        // Direct file URL if hosted
        hostedFileUrl: null // Set to your hosted .xlsx URL if applicable
    };

    excelBtn.addEventListener('click', function() {
        openExcel();
    });

    function openExcel() {
        // Add button click animation
        excelBtn.style.transform = 'scale(0.95)';
        setTimeout(() => {
            excelBtn.style.transform = '';
        }, 150);

        // Try multiple methods to open Excel
        
        // Method 1: Try to open local file via ms-excel protocol (works on Windows with Excel installed)
        const msExcelProtocol = `ms-excel:ofe|u|${window.location.origin}/${config.localFilePath}`;
        
        // Method 2: Direct file link
        const directLink = config.hostedFileUrl || config.localFilePath;
        
        // Method 3: Excel Online
        const excelOnline = config.onlineExcelUrl;

        // Attempt to open using ms-excel protocol first
        const protocolCheck = window.open(msExcelProtocol, '_self');
        
        // Fallback: if protocol doesn't work, try direct file or online
        setTimeout(() => {
            // If we're still on the page, the protocol probably didn't work
            // Try opening the file directly
            try {
                if (config.hostedFileUrl) {
                    window.open(config.hostedFileUrl, '_blank');
                    showToast('Opening Excel file...');
                } else {
                    // Fallback to Excel Online
                    window.open(excelOnline, '_blank');
                    showToast('Opening Excel Online...');
                }
            } catch (error) {
                showToast('Please download the file or use Excel Online');
            }
        }, 500);
    }

    function showToast(message) {
        toastMessage.textContent = message;
        toast.classList.add('show');
        
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    // Add keyboard support
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && document.activeElement === excelBtn) {
            openExcel();
        }
    });

    // Add ripple effect on button click
    excelBtn.addEventListener('click', function(e) {
        const ripple = document.createElement('span');
        const rect = this.getBoundingClientRect();
        const size = Math.max(rect.width, rect.height);
        const x = e.clientX - rect.left - size / 2;
        const y = e.clientY - rect.top - size / 2;
        
        ripple.style.cssText = `
            position: absolute;
            width: ${size}px;
            height: ${size}px;
            left: ${x}px;
            top: ${y}px;
            background: rgba(255, 255, 255, 0.3);
            border-radius: 50%;
            transform: scale(0);
            animation: ripple 0.6s ease-out;
            pointer-events: none;
        `;
        
        this.appendChild(ripple);
        
        setTimeout(() => ripple.remove(), 600);
    });

    // Add ripple animation
    const style = document.createElement('style');
    style.textContent = `
        @keyframes ripple {
            to {
                transform: scale(2);
                opacity: 0;
            }
        }
    `;
    document.head.appendChild(style);
});
