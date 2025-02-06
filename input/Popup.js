document.getElementById("start").addEventListener("click", async () => {
    const teamNames = await getTeamDataFromSheet();
    for (const name of teamNames) {
      await addTeamData(name);
    }
    alert("Proses input selesai!");
  });
  
  async function getTeamDataFromSheet() {
    const url = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec";
    const response = await fetch(url);
    const data = await response.json();
    return data;
  }
  
  async function addTeamData(name) {
    // Klik tombol Add New
    const addButton = document.querySelector("button:contains('Add New')");
    if (addButton) addButton.click();
  
    // Tunggu input form muncul
    await new Promise(r => setTimeout(r, 1000));
  
    // Isi input dengan nama tim
    const input = document.querySelector("input#input-650");
    if (input) {
      input.value = name;
      input.dispatchEvent(new Event('input', { bubbles: true }));
    }
  
    // Klik tombol Save
    const saveButton = document.querySelector("button:contains('Save')");
    if (saveButton) saveButton.click();
  
    // Tunggu selesai
    await new Promise(r => setTimeout(r, 1000));
  }  