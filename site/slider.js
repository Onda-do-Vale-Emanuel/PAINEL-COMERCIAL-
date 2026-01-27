document.addEventListener("DOMContentLoaded", () => {

  const slides = document.querySelectorAll(".slide");
  let index = 0;

  function trocar() {
    slides.forEach(s => s.classList.remove("active"));
    slides[index].classList.add("active");
    index = (index + 1) % slides.length;
  }

  trocar();
  setInterval(trocar, 5000);

});
