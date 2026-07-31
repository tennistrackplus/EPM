Office.onReady(() => {
  document.getElementById("closeBtn").onclick = () => {
    Office.addin.hideAsTaskpane();
  };
});
