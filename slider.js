let slideAtual = 0;
const slides = document.querySelectorAll(".slide");

function trocarSlide() {
  slides[slideAtual].classList.remove("ativo");
  slideAtual = (slideAtual + 1) % slides.length;
  slides[slideAtual].classList.add("ativo");
}

setInterval(trocarSlide, 15000); // 15 segundos
