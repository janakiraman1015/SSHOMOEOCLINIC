// Toggle Hamburger Menu
function toggleMenu(){
  const nav = document.getElementById('navLinks');
  const toggle = document.querySelector('.menu-toggle');
  nav.classList.toggle('active');
  toggle.classList.toggle('active');
}

// Close menu on link click
document.querySelectorAll('.nav-links a').forEach(link => {
  link.addEventListener('click', () => {
    document.getElementById('navLinks').classList.remove('active');
    document.querySelector('.menu-toggle').classList.remove('active');
  });
});