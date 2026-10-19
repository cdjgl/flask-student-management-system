function toggleSubmenu(menuId) {
    const submenu = document.getElementById(menuId);
    const menuItems = document.querySelectorAll('.menu-item.has-sub');
    
    menuItems.forEach(item => {
        const itemMenuId = item.getAttribute('onclick')?.match(/'([^']+)'/)?.[1];
        if (itemMenuId === menuId) {
            item.classList.toggle('open');
        }
    });
    
    if (submenu) {
        if (submenu.style.display === 'block') {
            submenu.style.display = 'none';
        } else {
            submenu.style.display = 'block';
        }
    }
}

function confirmDelete(message) {
    return confirm(message || '确定要删除吗？此操作不可恢复。');
}

function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
