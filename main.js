document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("callApiBtn");
  const output = document.getElementById("output");

  // 將此改為您在 Apps Script 部署後得到的 Web App URL
  // 例如： "https://script.google.com/macros/s/XXXXXXXXX/exec"
  const GAS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzl1OuvDsoj-ps73R8zQ_JyAWrSF7u4rSDoBL8ZoJIKsykdnD6K6dIBSpZf7YCsANR1dA/exec";

  btn.addEventListener("click", () => {
    // 例如要帶上參數 name
    const url = GAS_WEBAPP_URL + "?name=MichaelFromGitHub";

    // 使用 fetch 呼叫後端
    fetch(url)
      .then(response => {
        if (!response.ok) {
          throw new Error("HTTP error, status = " + response.status);
        }
        return response.json();
      })
      .then(data => {
        // data 即後端回傳的 JSON
        console.log("後端回傳:", data);
        output.textContent = JSON.stringify(data, null, 2);
      })
      .catch(err => {
        console.error("呼叫API失敗:", err);
        output.textContent = "呼叫API失敗: " + err.message;
      });
  });
});
