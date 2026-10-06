let questions = [];

async function loadExcel() {
  const file = await fetch("questions.xlsx");
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  questions = XLSX.utils.sheet_to_json(sheet);
  newQuestion();
}

function newQuestion() {
  const q = questions[Math.floor(Math.random() * questions.length)];

  document.getElementById("questionBox").innerText = q.question;

  const optionsBox = document.getElementById("optionsBox");
  optionsBox.innerHTML = "";

  ["A","B","C","D"].forEach(letter => {
    const btn = document.createElement("button");
    btn.classList.add("optionBtn");

    // 建立圖片
    if (q[letter + "_img"]) {
      const img = document.createElement("img");
      img.src = q[letter + "_img"];
      img.classList.add("optionImg");
      btn.appendChild(img);
    }

    // 建立文字（如果有）
    if (q[letter + "_text"]) {
      const text = document.createElement("div");
      text.innerText = q[letter + "_text"];
      btn.appendChild(text);
    }

    btn.onclick = () => {
      if (letter === q.answer) {
        btn.classList.add("correct");
      } else {
        btn.classList.add("wrong");
      }
    };

    optionsBox.appendChild(btn);
  });
}

document.getElementById("nextBtn").onclick = newQuestion;

loadExcel();
