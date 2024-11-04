let userName = "Usuário";

document.getElementById('importBtn').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert("Por favor, selecione um arquivo XLSX.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Seleciona a primeira planilha
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Converte os dados da planilha em um array
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Extrai a coluna A (índice 0) e coluna C (índice 2) a partir da linha 11 (índice 10)
        const extractedData = jsonData.slice(10).map(row => {
            const numberA = row[0]; // Dado da coluna A
            const dataC = row[2]; // Dado da coluna C
            return [numberA, dataC]; // Retorna como um array para colocar no novo XLSX
        }).filter(item => item[0] !== undefined || item[1] !== undefined); // Remove valores indefinidos

        // Exibe os dados extraídos
        document.getElementById('output').innerHTML = extractedData.length > 0 
            ? extractedData.map(data => `<div></div>`).join('') 
            : '<div>Nenhum dado encontrado nas colunas A e C.</div>';
        
        // Mostra o botão de exportação
        const exportBtn = document.getElementById('exportBtn');
        exportBtn.style.display = 'block';
        exportBtn.onclick = () => exportToXLSX(extractedData);
    };

    reader.readAsArrayBuffer(file);
});

function exportToXLSX(data) {
    // Formata a data atual
    const today = new Date();
    const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
    
    // Nome do arquivo
    const fileName = `Romaneio_${userName}_${formattedDate}.xlsx`;
    
    // Cria uma nova planilha
    const ws = XLSX.utils.aoa_to_sheet(data); // Coloca os dados na nova planilha
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dados");

    // Salva o arquivo
    XLSX.writeFile(wb, fileName);
}

document.getElementById('nameModal').style.display = 'block';

document.getElementById('submitNameBtn').addEventListener('click', function() {
    const userNameInput = document.getElementById('userNameInput');
    let userName = userNameInput.value || "Usuário"; // Pega o valor do input
    userNameInput.value = ""; // Limpa o campo após a entrada
    document.getElementById('nameModal').style.display = 'none'; // Fecha o modal
});


