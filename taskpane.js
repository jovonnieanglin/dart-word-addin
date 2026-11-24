(() => {
  const BYNDER_CLIENT_ID = "156b56aa-3eea-4520-8675-2d743673abe8";
  const BYNDER_AUTHORIZE_URL =
    "https://medialibrary.dart.ky/v6/authentication/oauth2/auth";

  const REDIRECT_URI = "https://jovonnieanglin.github.io";
  const SCOPE = "asset:read";

  const status = document.querySelector(".status");
  const title = document.getElementById("status-title");
  const detail = document.getElementById("status-detail");
  const insertButton = document.getElementById("insert-button");
  const result = document.getElementById("result");

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
      status.classList.add("error");
      title.textContent = "Open this add-in in Word";
      detail.textContent = "This test requires Microsoft Word.";
      return;
    }

    status.classList.add("ready");
    title.textContent = "Custom task pane is running";
    detail.textContent = "Connected to Microsoft Word.";

    insertButton.disabled = false;

    addBynderLogin();
  });

  insertButton.addEventListener("click", async () => {
    insertButton.disabled = true;
    result.textContent = "Writing to the document...";

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        body.insertParagraph(
          "DART Media Library add-in test successful.",
          Word.InsertLocation.end
        );

        await context.sync();
      });

      result.textContent =
        "Success: the add-in wrote to Word.";
    } catch (error) {
      result.textContent =
        "Test failed: " + (error.message || String(error));
    } finally {
      insertButton.disabled = false;
    }
  });

  function addBynderLogin() {
    const card = document.createElement("section");
    card.className = "card";

    const heading = document.createElement("h2");
    heading.textContent = "Bynder";

    const description = document.createElement("p");
    description.textContent =
      "Connect to the DART Media Library using your Bynder account.";

    const connectionStatus = document.createElement("p");
    connectionStatus.textContent = "Status: Not connected";

    const button = document.createElement("button");
    button.type = "button";
    button.textContent = "Sign in to Bynder";

    button.addEventListener("click", () => {
      startBynderAuthorization();
    });

    card.appendChild(heading);
    card.appendChild(description);
    card.appendChild(connectionStatus);
    card.appendChild(button);

    document.querySelector(".shell").appendChild(card);
  }

  function startBynderAuthorization() {
    const state = crypto.randomUUID();

    sessionStorage.setItem("bynder_oauth_state", state);

    const params = new URLSearchParams({
      client_id: BYNDER_CLIENT_ID,
      scope: SCOPE,
      redirect_uri: REDIRECT_URI,
      response_type: "code",
      state: state
    });

    const authorizationUrl =
      `${BYNDER_AUTHORIZE_URL}?${params.toString()}`;

    window.open(
      authorizationUrl,
      "_blank",
      "noopener,noreferrer"
    );
  }
})();
