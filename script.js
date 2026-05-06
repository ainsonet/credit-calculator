function clean(str) {
    return str.replace(/[^\d.]/g, '');
}

function pretty(num, suffix) {
    if (isNaN(num) || num === '' || num === null) return '—';
    let intNum = Math.round(num).toString();
    let intPart = intNum.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return intPart + ' ' + suffix;
}

function calculate() {
    document.getElementById('amount-error').textContent = '';
    document.getElementById('rate-error').textContent = '';
    document.getElementById('term-error').textContent = '';

    let amountVal = clean(document.getElementById('amount').value);
    let rateVal = clean(document.getElementById('rate').value);
    let termVal = clean(document.getElementById('term').value);

    let amount = parseFloat(amountVal);
    let rate = parseFloat(rateVal);
    let termYears = parseFloat(termVal);

    if (isNaN(amount) || amount <= 0) {
        document.getElementById('amount-error').textContent = 'Введите положительное число';
        return;
    }
    if (isNaN(rate) || rate < 0) {
        document.getElementById('rate-error').textContent = 'Введите неотрицательное число';
        return;
    }
    if (isNaN(termYears) || termYears <= 0) {
        document.getElementById('term-error').textContent = 'Введите положительное число';
        return;
    }

    let months = Math.round(termYears * 12);
    let monthlyRate = rate / 100 / 12;
    let monthlyPayment;
    if (monthlyRate === 0) {
        monthlyPayment = amount / months;
    } else {
        let pow = Math.pow(1 + monthlyRate, months);
        monthlyPayment = amount * monthlyRate * pow / (pow - 1);
    }
    let total = monthlyPayment * months;
    let overpay = total - amount;

    document.getElementById('monthly-payment').textContent = pretty(monthlyPayment, '₽');
    document.getElementById('loan-amount').textContent = pretty(amount, '₽');
    document.getElementById('overpayment').textContent = pretty(overpay, '₽');
    document.getElementById('total-amount').textContent = pretty(total, '₽');
}

window.addEventListener('DOMContentLoaded', function() {
    const amount = document.getElementById('amount');
    const rate = document.getElementById('rate');
    const term = document.getElementById('term');

    function inputHandler(e) {
        let field = e.target;
        let regex = /[^\d.]/g;
        if (field.id === 'amount' || field.id === 'term') {
            regex = /[^\d]/g;
        }
        let filtered = field.value.replace(regex, '');
        field.value = filtered;
        calculate();
    }

    function focusHandler(e) {
        let val = e.target.value.replace(/[^\d.]/g, '');
        e.target.value = val;
    }

    function blurHandler(e) {
        let field = e.target;
        let val = clean(field.value);
        let suffix = '';
        if (field.id === 'amount') suffix = '₽';
        else if (field.id === 'rate') suffix = '%';
        else if (field.id === 'term') suffix = 'лет';

        if (val) {
            field.value = pretty(parseFloat(val), suffix);
        } else {
            field.value = '';
        }
        calculate();
    }

    amount.addEventListener('input', inputHandler);
    rate.addEventListener('input', inputHandler);
    term.addEventListener('input', inputHandler);

    amount.addEventListener('focus', focusHandler);
    rate.addEventListener('focus', focusHandler);
    term.addEventListener('focus', focusHandler);

    amount.addEventListener('blur', blurHandler);
    rate.addEventListener('blur', blurHandler);
    term.addEventListener('blur', blurHandler);

    amount.value = '300 000 ₽';
    rate.value = '17 %';
    term.value = '5 лет';
    calculate();
});

function downloadCSV() {
    const rows = [
        ['Параметр', 'Значение'],
        ['Ежемесячный платеж', document.getElementById('monthly-payment').textContent],
        ['Сумма кредита', document.getElementById('loan-amount').textContent],
        ['Переплата', document.getElementById('overpayment').textContent],
        ['Общая сумма займа', document.getElementById('total-amount').textContent],
        ['Ставка', document.getElementById('rate').value],
        ['Срок', document.getElementById('term').value]
    ];

    let csv = '\uFEFF';
    rows.forEach(row => {
        csv += row.join(';') + '\n';
    });

    const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'расчёт_кредита.csv';
    a.click();
    URL.revokeObjectURL(url);
}

function saveCalculation() {
    const record = {
        date: new Date().toLocaleString('ru-RU'),
        amount: document.getElementById('amount').value,
        rate: document.getElementById('rate').value,
        term: document.getElementById('term').value,
        monthlyPayment: document.getElementById('monthly-payment').textContent,
        loanAmount: document.getElementById('loan-amount').textContent,
        overpayment: document.getElementById('overpayment').textContent,
        totalAmount: document.getElementById('total-amount').textContent
    };

    let history = JSON.parse(localStorage.getItem('calcHistory') || '[]');
    history.push(record);
    localStorage.setItem('calcHistory', JSON.stringify(history));
    alert('Расчёт сохранён!');
}

function downloadHistory() {
    const history = JSON.parse(localStorage.getItem('calcHistory') || '[]');
    const headers = ['Дата', 'Сумма кредита', 'Ставка', 'Срок', 'Ежемесячный платёж', 'Переплата', 'Общая сумма займа'];

    const data = [headers];
    history.forEach(item => {
        data.push([
            item.date,
            item.amount,
            item.rate,
            item.term,
            item.monthlyPayment,
            item.overpayment,
            item.totalAmount
        ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);

    ws['!cols'] = [
        {wch: 20},
        {wch: 22},
        {wch: 10},
        {wch: 10},
        {wch: 22},
        {wch: 20},
        {wch: 22}
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'История расчётов');
    XLSX.writeFile(wb, 'История Расчётов.xlsx');
}

function clearHistory() {
    if (confirm('Вы уверены, что хотите удалить всю историю расчётов?')) {
        localStorage.removeItem('calcHistory');
        alert('История очищена.');
    }
}