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
    
    if(latexInput.trim() === "") {
        previewDiv.innerHTML = '<p style="font-size: 12px; color: #666;">Preview aparecerá aqui</p>';
        if (errorAlert) errorAlert.style.display = "none";
        if (insertBtn) insertBtn.disabled = false;
        return;
    }
    
    try {
        katex.render(latexInput, previewDiv, {
            throwOnError: true,
            displayMode: true
        });
        
        if (errorAlert) errorAlert.style.display = "none";
        if (insertBtn) insertBtn.disabled = false;
        
    } catch (e) {
        if (errorAlert) errorAlert.style.display = "flex";
        if (insertBtn) insertBtn.disabled = true;
        
        if (errorMessage) {
            var cleanMessage = e.message.replace("KaTeX parse error: ", "");
            errorMessage.innerText = cleanMessage;
        }
        
        katex.render(latexInput, previewDiv, {
            throwOnError: false, 
            displayMode: true
        });
    }
}

function insertLatex() {
    var latexInput = document.getElementById("latexInput").value;
    if (!latexInput) return;

    try {
        var katexHtml = katex.renderToString(latexInput, {
            throwOnError: true, 
            displayMode: true,
            output: 'mathml'
        });

        var mathMlMatch = katexHtml.match(/<math[^>]*>[\s\S]*<\/math>/i);
        
        if (mathMlMatch) {
            var pureMathMl = mathMlMatch[0];
            
            // Garante o namespace XML
            if (!pureMathMl.includes("xmlns")) {
                pureMathMl = pureMathMl.replace("<math", '<math xmlns="http://www.w3.org/1998/Math/MathML"');
            }

            // Pacote OOXML para forçar o Word a converter para OMML nativo
            var ooxmlPayload = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
              <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                <pkg:xmlData>
                  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                  </Relationships>
                </pkg:xmlData>
              </pkg:part>
              <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                <pkg:xmlData>
                  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:body>
                      <w:altChunk r:id="altChunkId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                    </w:body>
                  </w:document>
                </pkg:xmlData>
              </pkg:part>
              <pkg:part pkg:name="/word/math.xml" pkg:contentType="application/mathml+xml">
                <pkg:xmlData>
                  ${pureMathMl}
                </pkg:xmlData>
              </pkg:part>
              <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                <pkg:xmlData>
                  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="altChunkId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="/word/math.xml"/>
                  </Relationships>
                </pkg:xmlData>
              </pkg:part>
            </pkg:package>`;

            Office.context.document.setSelectedDataAsync(
                ooxmlPayload,
                { coercionType: Office.CoercionType.Ooxml },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Erro ao inserir OOXML: " + asyncResult.error.message);
                    }
                }
            );
        }
    } catch (err) {
        console.error("Erro na injeção:", err);
    }
}
