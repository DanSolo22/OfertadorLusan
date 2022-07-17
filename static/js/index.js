let submenu = false;

console.log(submenu)

document.getElementById('close').onmousedown = function (e) {
    e.preventDefault();
    document.getElementById('msg').style.display = 'none';
    return false;
};

document.getElementById('4').onclick = function (e) {
    e.preventDefault();
    if (submenu) {
        document.getElementById('4').style.marginTop = '1.4vh';
        document.getElementById('drop').style.display = 'none';
        submenu = false;
    }
    else {
        document.getElementById('4').style.marginTop = '6.83vh';
        document.getElementById('drop').style.display = 'block';
        submenu = true;
    }
    return false;
}

function redirect(url) {
    window.location = url;
}

