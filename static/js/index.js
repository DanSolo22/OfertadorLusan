console.log(document.getElementById('close'))
console.log(document.getElementById('msg'))

document.getElementById('close').onmousedown = function(e) {
  e.preventDefault();
  document.getElementById('msg').style.display = 'none';
  return false;
};