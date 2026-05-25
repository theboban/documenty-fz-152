document.addEventListener('DOMContentLoaded', function() {
    const docTypeSelect = document.getElementById('doc_type');
    const newsletterSection = document.getElementById('newsletter_section');
    
    // Показываем секцию email для согласия на рассылку
    docTypeSelect.addEventListener('change', function() {
        if (this.value === 'consent_newsletter') {
            newsletterSection.style.display = 'block';
        } else {
            newsletterSection.style.display = 'none';
        }
    });
    
    // Валидация формы перед отправкой
    const form = document.getElementById('docForm');
    form.addEventListener('submit', function(e) {
        const docType = docTypeSelect.value;
        
        if (!docType) {
            e.preventDefault();
            alert('Пожалуйста, выберите тип документа');
            docTypeSelect.focus();
            return false;
        }
        
        // Проверка обязательных полей
        const requiredFields = ['org_name', 'org_address', 'org_email', 'site_url'];
        let isValid = true;
        
        requiredFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            if (field && !field.value.trim()) {
                isValid = false;
                field.style.borderColor = '#e74c3c';
                
                // Убираем подсветку ошибки при вводе
                field.addEventListener('input', function() {
                    this.style.borderColor = '#e0e0e0';
                }, { once: true });
            }
        });
        
        if (!isValid) {
            e.preventDefault();
            alert('Пожалуйста, заполните все обязательные поля (отмечены *)');
            return false;
        }
        
        // Показываем индикатор загрузки
        const submitBtn = document.querySelector('.btn-submit');
        submitBtn.innerHTML = '⏳ Генерация документа...';
        submitBtn.disabled = true;
        submitBtn.style.opacity = '0.7';
    });
    
    // Автоматическая установка текущей даты
    const dateInput = document.getElementById('date');
    if (!dateInput.value) {
        const now = new Date();
        const options = { day: 'numeric', month: 'long' };
        dateInput.value = now.toLocaleDateString('ru-RU', options);
    }
    
    // Плавная прокрутка к секциям при клике на заголовки
    document.querySelectorAll('.form-section h2').forEach(heading => {
        heading.style.cursor = 'pointer';
        heading.addEventListener('click', function() {
            this.parentElement.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        });
    });
    
    // Подсветка checkbox при выборе
    document.querySelectorAll('.checkbox-item input[type="checkbox"]').forEach(checkbox => {
        checkbox.addEventListener('change', function() {
            if (this.checked) {
                this.parentElement.style.background = '#eef1ff';
                this.parentElement.style.borderColor = '#667eea';
            } else {
                this.parentElement.style.background = '#f8f9fa';
                this.parentElement.style.borderColor = 'transparent';
            }
        });
        
        // Инициализация состояния
        if (checkbox.checked) {
            checkbox.parentElement.style.background = '#eef1ff';
            checkbox.parentElement.style.borderColor = '#667eea';
        }
    });
    
    console.log('Генератор документов 152-ФЗ готов к работе!');
});
