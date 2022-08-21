let submenu = false;

console.log(submenu)

document.getElementById('close').onmousedown = function (e) {
    e.preventDefault();
    document.getElementById('msg').style.display = 'none';
    return false;
};

/*document.getElementById('4').onmouseover = function (e) {
    e.preventDefault();
    document.getElementById('4').style.marginTop = '11.3vh';
    document.getElementById('drop').style.display = 'block';
}

document.getElementById('4').onmouseout = function (e) {
    e.preventDefault();
    document.getElementById('4').style.marginTop = '0';
    document.getElementById('drop').style.display = 'none';
}*/

function redirect(url) {
    window.location = url;
}

