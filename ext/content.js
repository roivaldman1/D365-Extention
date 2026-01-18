chrome.runtime.onMessage.addListener((msg) => {
  if (msg.action === "ribbondebug")  ribbondebug(msg); 


});

const ribbondebug = (msg) => {
    console.log(Xrm)
}