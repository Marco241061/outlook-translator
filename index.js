Office.onReady(() => {
    // Inizializza l'add-in quando Office è pronto
    document.getElementById('translateButton').onclick = translateText;
});

let lastText = '';
const targetLanguage = 'en'; // Lingua di destinazione (inglese in questo esempio)

async function translateText() {
    try {
        // Ottiene il testo corrente dall'email
        Office.context.mailbox.item.body.getAsync('text', async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const currentText = result.value;
                
                // Traduci solo il nuovo testo
                if (currentText !== lastText) {
                    const translatedText = await translateWithAPI(currentText);
                    
                    // Aggiorna il corpo dell'email con la traduzione
                    Office.context.mailbox.item.body.setAsync(
                        translatedText,
                        { coercionType: 'text' },
                        (result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                lastText = translatedText;
                            }
                        }
                    );
                }
            }
        });
    } catch (error) {
        console.error('Errore durante la traduzione:', error);
    }
}

async function translateWithAPI(text) {
    // Qui dovrai implementare la chiamata al servizio di traduzione
    // Per esempio, usando Microsoft Translator API
    const endpoint = 'https://api.cognitive.microsofttranslator.com/translate';
    const subscriptionKey = 'TUA_CHIAVE_API';

    const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Ocp-Apim-Subscription-Key': subscriptionKey,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify([{
            text: text
        }])
    });

    const translation = await response.json();
    return translation[0].translations[0].text;
} 