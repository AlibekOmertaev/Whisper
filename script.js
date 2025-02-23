document.getElementById("startDictation").addEventListener("click", function() {
    let audioFile = document.createElement('input');
    audioFile.type = 'file';
    audioFile.accept = 'audio/*';
    audioFile.onchange = function () {
        let file = audioFile.files[0];

        let formData = new FormData();
        formData.append("file", file);
        formData.append("model", "whisper-1");
        formData.append("language", "ru");

        fetch("https://api.openai.com/v1/audio/transcriptions", {
            method: "POST",
            headers: {
                "Authorization": "Bearer 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"
            },
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            let text = data.text;

            // Вставляем текст в Word
            Office.context.document.setSelectedDataAsync(text, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Ошибка вставки текста: " + asyncResult.error.message);
                }
            });
        })
        .catch(error => console.error("Ошибка запроса:", error));
    };

    audioFile.click();
});