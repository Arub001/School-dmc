document.getElementById("uploadExcel").addEventListener("change", handleFile);
document.getElementById("printAll").addEventListener("click", () => window.print());

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const students = XLSX.utils.sheet_to_json(sheet);

    generateCards(students);
  };
  reader.readAsArrayBuffer(file);
}

function generateCards(students) {
  const container = document.getElementById("resultCards");
  container.innerHTML = "";

  students.forEach((student) => {
    const card = document.createElement("div");
    card.className = "result-card";

    card.innerHTML = `
      <h2>Result Card</h2>
      <div class="student-info">
        <div><strong>Name:</strong> ${student["student name"] || "N/A"}</div>
        <div><strong>Father's Name:</strong> ${student["father name"] || "N/A"}</div>
      </div>
      <table>
        <thead>
          <tr>
            <th>Subject</th>
            <th>Marks</th>
          </tr>
        </thead>
        <tbody>
          <tr><td>Chemistry</td><td>${student["chemistry"] || "N/A"}</td></tr>
          <tr><td>Physics</td><td>${student["physics"] || "N/A"}</td></tr>
          <tr><td>Math</td><td>${student["math"] || "N/A"}</td></tr>
          <tr><td>History</td><td>${student["history"] || "N/A"}</td></tr>
        </tbody>
      </table>
      <div class="footer">
        <div><strong>Total Marks:</strong> ${student["total marks"] || "N/A"}</div>
        <div><strong>Obtained Marks:</strong> ${student["obtained marks"] || "N/A"}</div>
        <div><strong>Percentage:</strong> ${student["percentage"] || "N/A"}%</div>
        <div><strong>Grade:</strong> ${student["grade"] || "N/A"}</div>
      </div>
      <div class="signatures">
        <div>Class Teacher</div>
        <div>Parent</div>
        <div>Principal</div>
      </div>
    `;

    container.appendChild(card);
  });
}