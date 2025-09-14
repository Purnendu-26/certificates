// Template preview
document.getElementById("template").addEventListener("change", function (event) {
  const file = event.target.files[0];
  if (file) {
    const reader = new FileReader();
    reader.onload = function (e) {
      const previewImg = document.getElementById("templatePreview");
      previewImg.src = e.target.result;
      previewImg.style.display = "block";
    };
    reader.readAsDataURL(file);
  }
});

const uploadForm = document.getElementById("uploadForm");
const downloadBtn = document.getElementById("downloadBtn");
const statusMsg = document.getElementById("statusMsg");
const spinner = document.getElementById("spinner");

uploadForm.addEventListener("submit", function (e) {
  e.preventDefault();

  spinner.classList.remove("hidden");
  statusMsg.innerText = "⏳ Generating certificates, please wait...";
  downloadBtn.classList.add("hidden");

  const formData = new FormData(uploadForm);

  fetch("/upload", {
    method: "POST",
    body: formData,
  })
    .then((res) => res.json())
    .then((data) => {
      spinner.classList.add("hidden");

      if (data.message) {
        statusMsg.innerText = "✅ Certificates generated successfully!";
        downloadBtn.classList.remove("hidden");
      } else {
        statusMsg.innerText = "❌ Error: " + (data.error || "Unknown error");
      }
    })
    .catch(() => {
      spinner.classList.add("hidden");
      statusMsg.innerText = "⚠️ Server error. Try again.";
    });
});
