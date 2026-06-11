(() => {
  const RENDER_EVENT = "streamlit:render";
  const COMPONENT_READY = "streamlit:componentReady";
  const SET_COMPONENT_VALUE = "streamlit:setComponentValue";
  const SET_FRAME_HEIGHT = "streamlit:setFrameHeight";
  const API_VERSION = 1;

  const eventTarget = new EventTarget();

  function sendMessage(message) {
    if (window.parent) {
      window.parent.postMessage({ isStreamlitMessage: true, ...message }, "*");
    }
  }

  function handleMessage(event) {
    if (!event || !event.data || event.data.type !== RENDER_EVENT) {
      return;
    }
    eventTarget.dispatchEvent(
      new CustomEvent(RENDER_EVENT, { detail: event.data })
    );
  }

  window.addEventListener("message", handleMessage);

  window.Streamlit = {
    RENDER_EVENT,
    events: eventTarget,
    setComponentReady: () => {
      sendMessage({ type: COMPONENT_READY, apiVersion: API_VERSION });
    },
    setFrameHeight: (height) => {
      const frameHeight =
        typeof height === "number" && height > 0
          ? height
          : document.documentElement.scrollHeight || document.body.scrollHeight;
      sendMessage({ type: SET_FRAME_HEIGHT, height: frameHeight });
    },
    setComponentValue: (value) => {
      sendMessage({ type: SET_COMPONENT_VALUE, value: value });
    },
  };
})();
