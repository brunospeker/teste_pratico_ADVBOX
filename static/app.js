function reconheceInput() {
    const fileInput = document.getElementById('fileInput');
    if ('webkitdirectory' in fileInput) {
        fileInput.setAttribute('webkitdirectory', '');
        fileInput.setAttribute('multiple', '');
        document.getElementById('info').innerText = "2º - Selecione a pasta que contém os arquivos do rar e clique no botão enviar.";
    } else {
        fileInput.removeAttribute('webkitdirectory');
        fileInput.setAttribute('multiple', '');
        document.getElementById('info').innerText = "2º - Selecione todos os arquivos contidos dentro do rar e clique no botão enviar.";
    }
}

document.addEventListener('DOMContentLoaded', reconheceInput);

//Envio dos arquivos para a api
const form = document.getElementById('uploadForm');
const responseMessage = document.getElementById('responseMessage');

form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const fileInput = document.getElementById('fileInput');
    const files = fileInput.files;

    if (files.length === 0) {
        responseMessage.textContent = 'Nenhum arquivo selecionado.';
        responseMessage.style.color = 'red';
        return;
    }

    const formData = new FormData();
    for (const file of files) {
        formData.append('files', file);
    }

    document.querySelector('button[type="submit"]').disabled = true;
    document.getElementById('linkProcessos').classList.add('disabled');
    document.getElementById('linkClientes').classList.add('disabled');
    document.getElementById('linkProcessos').setAttribute('aria-disabled', 'true');
    document.getElementById('linkClientes').setAttribute('aria-disabled', 'true');

    try {
        const response = await fetch('/upload_file', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Erro ao enviar arquivos');
        }

        const result = await response.json();
        responseMessage.textContent = result['message'];
        responseMessage.style.color = 'green';

        document.getElementById('linkProcessos').classList.remove('disabled');
        document.getElementById('linkClientes').classList.remove('disabled');
        document.getElementById('linkProcessos').removeAttribute('aria-disabled');
        document.getElementById('linkClientes').removeAttribute('aria-disabled');
    } catch (error) {
        responseMessage.textContent = 'Erro: ' + error.message;
        responseMessage.style.color = 'red';
    } finally {
        document.querySelector('button[type="submit"]').disabled = false;
    }
});