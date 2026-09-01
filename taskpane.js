// Aguarda o Office.js carregar completamente antes de liberar as interações
Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("insertBtn").onclick = insertLatex;
        document.getElementById("latexInput").addEventListener("input", updatePreview);
    }
});

// Atualiza o preview dinamicamente (DRY: reaproveita a lógica de montagem da URL)
function getLatexImageUrl(latexText) {
    // dpi alto para boa resolução no Word
    return "https://latex.codecogs.com/png.latex?\\dpi{300}\\huge " + encodeURIComponent(latexText);
}

function updatePreview() {
    var latexInput = document.getElementById("latexInput").value;
    var previewDiv = document.getElementById("preview");
    
    if(latexInput.trim() === "") {
        previewDiv.innerHTML = '<p style="font-size: 12px; color: #666;">Preview aparecerá aqui</p>';
        return;
    }
    
    previewDiv.innerHTML = "<img src='" + getLatexImageUrl(latexInput) + "' alt='preview' style='max-width:100%;' />";
}

function insertLatex() {
    var latexInput = document.getElementById("latexInput").value;
    if (!latexInput) return;

    var imgUrl = getLatexImageUrl(latexInput);
    var htmlContent = "<img src='" + imgUrl + "' alt='Equação LaTeX' />";

    // Insere o HTML (que carrega a imagem da equação) no local do cursor
    Office.context.document.setSelectedDataAsync(
        htmlContent,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Erro ao inserir: " + asyncResult.error.message);
            }
        }
    );
}