/* global Office */

Office.onReady(() => {
  Office.actions.associate('translateToFrysk', translateToFrysk);
});

async function translateToFrysk(event) {
  const item = Office.context.mailbox.item;

  try {
    const selected = await getSelectedData(item, Office.CoercionType.Text);

    if (!selected || !selected.trim()) {
      await notify(item, 'info', 'Selecteer eerst de te vertalen tekst.');
      event.completed();
      return;
    }

    const translation = await callFryskerAPI(selected);
    await setSelectedData(item, translation, Office.CoercionType.Text);
    await notify(item, 'info', 'Oersetting klear! ✓');

  } catch (err) {
    await notify(item, 'error', 'Vertaling mislukt: ' + err.message);
  }

  event.completed();
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

function getSelectedData(item, coercionType) {
  return new Promise((resolve, reject) => {
    item.getSelectedDataAsync(coercionType, result => {
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve(result.value)
        : reject(new Error(result.error.message));
    });
  });
}

function setSelectedData(item, content, coercionType) {
  return new Promise((resolve, reject) => {
    item.setSelectedDataAsync(content, { coercionType }, result => {
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
