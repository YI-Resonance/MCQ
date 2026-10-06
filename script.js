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

  // 題目欄位叫 char，不是 question
  document.getElementById("questionBox").innerText = q.char;

  const optionsBox = document.getElementById("optionsBox");
  optionsBox.innerHTML = "";

  // Excel 有 A、B、C、D、E 五個選項
  ["A","B","C","D","E"].forEach(letter => {
    const btn = document.createElement("button");
    btn.classList.add("optionBtn");

    const imgField = letter + "_img";
    const textField = letter + "_text";

    // 圖片
    if (q[imgField]) {
      const img = document.createElement("img");
      img.src = q[imgField];
      img.classList.add("optionImg");
      btn.appendChild(img);
    }

    // 文字
    if (q[textField]) {
      const text = document.createElement("div");
      text.innerText = q[textField];
      btn.appendChild(text);
    }

    // 答案判斷（Excel 的答案欄位叫 Ans）
    btn.onclick = () => {
      if (letter === q.Ans) {
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
