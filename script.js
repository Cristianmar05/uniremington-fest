// ============================================================
// SEMANA DE ANIVERSARIO UNIREMINGTON — SCRIPT
// ============================================================

const cards = document.querySelectorAll('.card');
const burgerBtn = document.getElementById('burgerBtn');
const navLinks = document.getElementById('navLinks');
const navbar = document.getElementById('navbar');

// Elementos del Modal
const modal = document.getElementById('photoModal');
const modalBackdrop = document.getElementById('modalBackdrop');
const modalClose = document.getElementById('modalClose');
const modalTitle = document.getElementById('modalTitle');
const modalDate = document.getElementById('modalDate');
const modalImgInicio = document.getElementById('modalImgInicio');
const modalImgFin = document.getElementById('modalImgFin');
const excelBtns = document.querySelectorAll('.excel-btn');

// ============================================================
// 1. MANEJO DE CLIC EN LA TARJETA (Abrir Modal)
// ============================================================
cards.forEach(card => {
  card.addEventListener('click', (e) => {
    // Si se hizo clic en el botón de Excel, no abrir el modal
    if (e.target.closest('.excel-btn')) return;

    // Obtener información de la tarjeta
    const title = card.dataset.title;
    const date = card.dataset.date;
    const imgInicio = card.dataset.imgInicio;
    const imgFin = card.dataset.imgFin;

    // Rellenar modal
    modalTitle.textContent = title;
    modalDate.textContent = date;
    modalImgInicio.src = imgInicio;
    modalImgFin.src = imgFin;

    // Mostrar modal
    modal.classList.add('active');
    document.body.style.overflow = 'hidden'; // Prevenir scroll de fondo
  });
});

// ============================================================
// 2. MANEJO DEL BOTÓN EXCEL (Detener Propagación)
// ============================================================
excelBtns.forEach(btn => {
  btn.addEventListener('click', (e) => {
    e.stopPropagation(); // Evita conflictos con el clic de la tarjeta
  });
});

// ============================================================
// 3. CERRAR MODAL
// ============================================================
function closeModal() {
  modal.classList.remove('active');
  document.body.style.overflow = ''; // Restaurar scroll
}

if (modalClose) modalClose.addEventListener('click', closeModal);
if (modalBackdrop) modalBackdrop.addEventListener('click', closeModal);

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && modal.classList.contains('active')) {
    closeModal();
  }
});

// ============================================================
// 4. NAVBAR SCROLL EFFECT
// ============================================================
window.addEventListener('scroll', () => {
  navbar.classList.toggle('navbar--scrolled', window.scrollY > 20);
}, { passive: true });

// ============================================================
// 5. MOBILE MENU (Burger)
// ============================================================
if (burgerBtn) {
  burgerBtn.addEventListener('click', () => navLinks.classList.toggle('open'));
}
navLinks.querySelectorAll('.navbar__link').forEach(link => {
  link.addEventListener('click', () => navLinks.classList.remove('open'));
});

// ============================================================
// 6. INTERSECTION OBSERVER — Animación de entrada de tarjetas
// ============================================================
const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry, i) => {
    if (entry.isIntersecting) {
      setTimeout(() => entry.target.classList.add('visible'), i * 120);
      observer.unobserve(entry.target);
    }
  });
}, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });

cards.forEach(card => observer.observe(card));

// ============================================================
// 7. SMOOTH SCROLL — Navegación suave a secciones
// ============================================================
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
  anchor.addEventListener('click', function (e) {
    const target = document.querySelector(this.getAttribute('href'));
    if (target) {
      e.preventDefault();
      const offset = navbar.offsetHeight + 16;
      const top = target.getBoundingClientRect().top + window.pageYOffset - offset;
      window.scrollTo({ top, behavior: 'smooth' });
    }
  });
});
