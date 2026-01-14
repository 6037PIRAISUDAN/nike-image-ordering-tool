let finalOutput = [];

document.getElementById("excelFile").addEventListener("change", handleExcel);
document.getElementById("submitBtn").addEventListener("click", exportExcel);

function handleExcel(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(evt.target.result, { type: "binary" });
    const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    renderGrid(data);
  };
  reader.readAsBinaryString(file);
}

function renderGrid(data) {
  const grid = document.getElementById("grid");
  grid.innerHTML = "";

  data.forEach(row => {
    const style = row["Style Code"];
    if (!style) return;

    const title = document.createElement("div");
    title.className = "style-title";
    title.innerText = style;
    grid.appendChild(title);

    const rowDiv = document.createElement("div");
    rowDiv.className = "grid-row";

    Object.keys(row).forEach(key => {
      if (key !== "Style Code" && row[key]) {
        const box = document.createElement("div");
        box.className = "image-box";
        box.draggable = true;

        const img = document.createElement("img");
        img.src = row[key];
        img.onerror = () => box.remove();

        const btn = document.createElement("button");
        btn.className = "remove";
        btn.innerText = "X";
        btn.onclick = () => box.remove();

        box.appendChild(btn);
        box.appendChild(img);
        rowDiv.appendChild(box);
      }
    });

    enableSmoothDrag(rowDiv);
    grid.appendChild(rowDiv);
  });
}

/* SMOOTH DRAG */
function enableSmoothDrag(container) {
  let dragged = null;

  container.querySelectorAll(".image-box").forEach(box => {
    box.addEventListener("dragstart", e => {
      dragged = box;
      box.style.opacity = "0.4";
    });

    box.addEventListener("dragend", () => {
      dragged.style.opacity = "1";
      dragged = null;
    });

    box.addEventListener("dragover", e => e.preventDefault());

    box.addEventListener("drop", () => {
      if (dragged && dragged !== box) {
        box.style.transition = "transform 0.3s ease";
        box.before(dragged);
      }
    });
  });
}

function exportExcel() {
  finalOutput = [];

  document.querySelectorAll(".style-title").forEach(title => {
    const style = title.innerText;
    const images = title.nextElementSibling.querySelectorAll("img");

    const row = { "Style Code": style };
    images.forEach((img, i) => row[`Image_${i + 1}`] = img.src);
    finalOutput.push(row);
  });

  const ws = XLSX.utils.json_to_sheet(finalOutput);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Final");
  XLSX.writeFile(wb, "output.xlsx");
}