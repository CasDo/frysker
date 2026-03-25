/* global Office, DOMParser */

Office.onReady(() => {
  Office.actions.associate('translateToFrysk', translateToFrysk);
});

async function translateToFrysk(event) {
  const item = Office.context.mailbox.item;

  try {
    const html = await getBody(item, Office.CoercionType.Html);
    const text = extractText(html);

    if (!text.trim()) {
      await notify(item, 'info', 'De e-mail bevat geen tekst om te vertalen.');
      event.completed();
      return;
    }

    const translation = await callFryskerAPI(text);
    const translatedHtml = rebuildHtml(html, translation);

    await setBody(item, translatedHtml, Office.CoercionType.Html);
    await notify(item, 'info', 'Oersetting klear! ✓');

  } catch (err) {
    await notify(item, 'error', 'Vertaling mislukt: ' + err.message);
  }

  event.completed();
}

// ------------------------------------------------------------

/** Haalt platte tekst uit HTML met alineastructuur bewaard */
function extractText(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
  doc.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6').forEach(el => {
    if (el.textContent.trim()) el.insertAdjacentText('afterend', '\n\n');
  });
  return (doc.body.innerText || '').replace(/\n{3,}/g, '\n\n').trim();
}

/** Bouwt nieuwe HTML op met de vertaalde tekst, bewaart e-mailopmaak */
function rebuildHtml(originalHtml, translatedText) {
  const doc = new DOMParser().parseFromString(originalHtml, 'text/html');

  const firstDiv = doc.body.querySelector('div[style]');
  const containerStyle = firstDiv
    ? firstDiv.getAttribute('style')
    : 'font-family: Calibri, sans-serif; font-size: 11pt;';

  const paragraphs = translatedText
    .split('\n\n')
    .filter(p => p.trim())
    .map(p => `<p>${p.replace(/\n/g, '<br>')}</p>`)
    .join('');

  doc.body.innerHTML = `<div style="${containerStyle}">${paragraphs}</div>`;
  return doc.documentElement.outerHTML;
}

// ------------------------------------------------------------

async function callFryskerAPI(text) {
  const response = await fetch('https://frisian.eu/languageapi_v8', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      operationName: 'TranslateFunction',
      variables: { text, lang: 'nl' },
      query: `query TranslateFunction($text: String!, $lang: LangType!) {
  translatetext(text: $text, lang: $lang) {
    translation
    __typename
  }
}`,
    }),
  });

  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const data = await response.json();
  const translation = data?.data?.translatetext?.translation;
  if (!translation) throw new Error('Geen vertaling ontvangen');
  return translation;
}

// ------------------------------------------------------------

function getBody(item, coercionType) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(coercionType, result => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error(result.error.message));
    });
  });
}

function setBody(item, content, coercionType) {
  return new Promise((resolve, reject) => {
    item.body.setAsync(content, { coercionType }, result => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve()
        : reject(new Error(result.error.message));
    });
  });
}

function notify(item, type, message) {
  return new Promise(resolve => {
    const msgType = type === 'error'
      ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
    item.notificationMessages.replaceAsync('frysker', {
      type: msgType,
      ...(type !== 'error' && { icon: 'Icon.16', persistent: false }),
      message,
    }, () => resolve());
  });
}
