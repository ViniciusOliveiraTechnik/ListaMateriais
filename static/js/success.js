const table = document.querySelector('table');
const tableData = JSON.parse(document.getElementById('table-data').textContent);

// Função para padronizar objetos como string JSON
const normalize = (obj) => JSON.stringify(obj, Object.keys(obj).sort());

table.addEventListener('click', (event) => {
    const deleteBtn = event.target.closest('a');
    if (deleteBtn) {
        const row = deleteBtn.closest('tr');
        if (row) {
            const content = row.querySelectorAll('td');
            const data = {
                description: content[0].textContent.trim(),
                spec: content[1].textContent.trim(),
                size: content[2].textContent.trim(),
                length: parseFloat(content[3].textContent.trim()),
                categorie: content[4].textContent.trim(),
            };

            // Normalizar objetos para comparação
            const dataString = normalize(data);

            // Encontrar o índice
            const index = tableData.findIndex(item => normalize(item) === dataString);

            // Remover a linha da tabela
            row.remove();

            // Opcional: Remover do array `tableData`
            if (index !== -1) {
                tableData.splice(index, 1);
            }
        }
    }
});

const downloadBtn = document.getElementById('download-btn');

downloadBtn.addEventListener('click', async () => {
    const fileInput = document.getElementById('file-name');

    // Obtém o nome do arquivo selecionado
    const fileName = fileInput.value;

    try {
        const response = await fetch('/download_wb/', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-CSRFToken': getCookie('csrftoken'),
            },
            body: JSON.stringify(tableData),
        });

        if (!response.ok) {
            alert(`Erro de execução: ${response.status}`)
            throw new Error(`Erro HTTP: ${response.status}`);
        }
        
        // Tratar a resposta como um Blob para download de arquivo binário
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        

        // Criar um link de download
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;  // Nome do arquivo

        document.body.appendChild(a);
        a.click();

        // Limpar o link após o download
        document.body.removeChild(a);

    } catch (error) {
        console.error('Erro na solicitação', error);
    }
});

// Função para obter o token CSRF
function getCookie(name) {
    const cookieValue = document.cookie
        .split('; ')
        .find(row => row.startsWith(name))
        ?.split('=')[1];
    return cookieValue || ''; 
}
