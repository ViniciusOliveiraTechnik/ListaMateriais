const fileInput = document.getElementById('files');
const fileListContainer = document.querySelector('.files-container');
const fileList = document.querySelector('.files-list');

const rangeInput = document.getElementById('percentual');
const rangeSpan = document.getElementById('percentual-value');

const resetBtn = document.getElementById('reset-btn');

resetBtn.addEventListener('click', () => {
    fileList.innerHTML = '';
    fileListContainer.style.display = 'none';
})

fileInput.addEventListener('change', () => {
    
    fileList.innerHTML = '';
    const files = fileInput.files;

    if(files.length > 0) {
        fileListContainer.style.display = 'block';

        Array.from(files).forEach(file => {
            const listItem = document.createElement('li');
            listItem.textContent = file.name;
            fileList.appendChild(listItem);
        });
    } else {
        fileListContainer.style.display = 'none';
    }

})

rangeInput.addEventListener('input', () => {
    rangeSpan.innerHTML = '';
    rangeSpan.innerHTML = `${rangeInput.value}%`;
})