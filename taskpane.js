(() => {
  const status = document.querySelector('.status');
  const title = document.getElementById('status-title');
  const detail = document.getElementById('status-detail');
  const button = document.getElementById('insert-button');
  const result = document.getElementById('result');

  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      status.classList.add('ready');
      title.textContent = 'Custom task pane is running';
      detail.textContent = 'Connected to Microsoft Word.';
      button.disabled = false;
    } else {
      status.classList.add('error');
      title.textContent = 'Open this add-in in Word';
      detail.textContent = 'The current Office host is not supported by this test.';
    }
  });

  button.addEventListener('click', async () => {
    button.disabled = true;
    result.textContent = 'Writing to the document...';
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.insertParagraph('DART Media Library add-in test successful.', Word.InsertLocation.end);
        await context.sync();
      });
      result.textContent = 'Success: the add-in wrote to Word.';
    } catch (error) {
      result.textContent = 'Test failed: ' + (error.message || String(error));
    } finally {
      button.disabled = false;
    }
  });
})();
