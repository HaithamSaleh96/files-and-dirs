
const apiUrl = 'http://localhost:3000/api';


function isExcelFile(fileName) {
  return fileName.endsWith('.xlsx');
}

// Read file
async function readFile() {
  const fileName = document.getElementById('readFileName').value;
  
  if (!isExcelFile(fileName)) {
    alert("Only Excel files are allowed.");
    return;
  }

  const response = await fetch(`${apiUrl}/read?fileName=${fileName}`);
  const data = await response.json();
  

  document.getElementById('readContent').textContent = JSON.stringify(data.content || data.error, null, 2);
}

// Write file
async function writeFile() {
  const fileName = document.getElementById('writeFileName').value;
  if (!isExcelFile(fileName)) {
    alert("Only Excel files are allowed.");
    return;
  }
  const content = document.getElementById('writeContent').value;

  
  const jsonContent = [[content]]; 

  await fetch(`${apiUrl}/write`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileName, content: jsonContent }) 
  });
}

// Append file
async function appendFile() {
  const fileName = document.getElementById('appendFileName').value;
  if (!isExcelFile(fileName)) {
    alert("Only Excel files are allowed.");
    return;
  }
  const content = document.getElementById('appendContent').value;

 
  const jsonContent = [[content]]; 

  await fetch(`${apiUrl}/append`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileName, content: jsonContent }) 
  });
}


// Delete file
async function deleteFile() {
  const fileName = document.getElementById('deleteFileName').value;

  if (!isExcelFile(fileName)) {
    alert("Only Excel files are allowed.");
    return;
  }

  await fetch(`${apiUrl}/delete?fileName=${fileName}`, { method: 'DELETE' });
}

// Rename file
async function renameFile() {
  const oldName = document.getElementById('oldFileName').value;
  const newName = document.getElementById('newFileName').value;

  if (!isExcelFile(oldName) || !isExcelFile(newName)) {
    alert("Only Excel files are allowed.");
    return;
  }

  await fetch(`${apiUrl}/rename`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ oldName, newName })
  });
}
