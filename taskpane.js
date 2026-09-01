Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("insertBtn").onclick = insertLatex;
        document.getElementById("latexInput").addEventListener("input", updatePreview);
    }
});

function updatePreview() {
    var latexInput = document.getElementById("latexInput").value;
    var previewDiv = document.getElementById("preview");
    var errorAlert = document.getElementById("errorAlert");
    var errorMessage = document.getElementById("errorMessage");
    var insertBtn = document.getElementById("insertBtn");

    if (latexInput.trim() === "") {
        previewDiv.innerHTML = '<p style="font-size: 12px; color: #666;">Preview aparecerá aqui</p>';
        errorAlert.style.display = "none";
        insertBtn.disabled = false;
        return;
    }

    try {
        katex.render(latexInput, previewDiv, { throwOnError: true, displayMode: true });
        errorAlert.style.display = "none";
        insertBtn.disabled = false;
    } catch (e) {
        errorAlert.style.display = "flex";
        insertBtn.disabled = true;
        errorMessage.innerText = e.message.replace("KaTeX parse error: ", "");
        katex.render(latexInput, previewDiv, { throwOnError: false, displayMode: true });
    }
}

function insertLatex() {
    var latexInput = document.getElementById("latexInput").value;
    var errorAlert = document.getElementById("errorAlert");
    var errorMessage = document.getElementById("errorMessage");
    
    if (!latexInput) return;

    try {
        var katexHtml = katex.renderToString(latexInput, { throwOnError: true, displayMode: true, output: 'mathml' });
        var mathMlMatch = katexHtml.match(/<math[^>]*>[\s\S]*<\/math>/i);

        if (mathMlMatch) {
            var pureMathMl = mathMlMatch[0];
            
            // Limpeza das tags auxiliares do KaTeX
            pureMathMl = pureMathMl.replace(/<semantics[^>]*>/gi, '');
            pureMathMl = pureMathMl.replace(/<\/semantics>/gi, '');
            pureMathMl = pureMathMl.replace(/<annotation[^>]*>[\s\S]*?<\/annotation>/gi, '');
            
            if (pureMathMl.indexOf("xmlns") === -1) {
                pureMathMl = pureMathMl.replace("<math", '<math xmlns="http://www.w3.org/1998/Math/MathML"');
            }

            Word.run(function (context) {
                var range = context.document.getSelection();
                
                // Roteia o MathML puro pelo conversor HTML do Word
                range.insertHtml(pureMathMl, Word.InsertLocation.replace);
                
                return context.sync().then(function () {
                    errorAlert.style.display = "none";
                });
            }).catch(function (error) {
                errorAlert.style.display = "flex";
                errorMessage.innerText = "Erro na injeção: " + error.message;
            });
        }
    } catch (err) {
        errorAlert.style.display = "flex";
        errorMessage.innerText = "Erro no Script: " + err.message;
    }
}
