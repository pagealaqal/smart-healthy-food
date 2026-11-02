/* Smart Healthy Food — تفاعلات الموقع */
var menuItems = [
  { name: 'وعاء السلمون المتوازن', price: 42, calories: 480 },
  { name: 'وعاء الدجاج الأخضر', price: 36, calories: 420 },
  { name: 'راب الحبوب الكاملة', price: 28, calories: 350 },
  { name: 'سموذي التوت', price: 19, calories: 210 },
  { name: 'كرات الطاقة', price: 16, calories: 160 }
];

function showToast(message) {
  var toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add('show');
  window.setTimeout(function () { toast.classList.remove('show'); }, 3200);
}

function showStatus(form, message, type) {
  var status = form.querySelector('.form-status');
  if (status) {
    status.textContent = message;
    status.className = 'form-status ' + type;
  }
}

function calculateOrderTotal() {
  var total = 0;
  var count = 0;
  var output = document.getElementById('calc-output');
  var rows = document.querySelectorAll('[data-calc-item]');
  for (var i = 0; i < rows.length; i++) {
    var quantity = parseInt(rows[i].querySelector('input').value, 10) || 0;
    var price = Number(rows[i].getAttribute('data-price'));
    total = total + (price * quantity);
    count = count + quantity;
  }
  var totalNode = document.getElementById('calc-total');
  var countNode = document.getElementById('calc-count');
  if (totalNode) totalNode.textContent = total.toFixed(2) + ' ر.س';
  if (countNode) countNode.textContent = count;
  if (output) {
    if (count === 0) {
      output.textContent = 'اختاري وجبتك المفضلة، وسيظهر الإجمالي هنا.';
    } else if (total >= 100) {
      output.textContent = 'رائع! طلبك يتجاوز 100 ريال — ' + count + ' صنفًا في السلة.';
    } else {
      output.textContent = 'تم احتساب ' + count + ' صنفًا. أضيفي المزيد حسب رغبتك.';
    }
  }
  return total;
}

function updateOrderSummary() {
  var select = document.getElementById('order-meal');
  var quantityField = document.getElementById('order-quantity');
  if (!select || !quantityField) return;
  var selected = select.options[select.selectedIndex];
  var price = Number(selected.getAttribute('data-price')) || 0;
  var quantity = parseInt(quantityField.value, 10) || 0;
  var subtotal = price * quantity;
  var delivery = subtotal > 0 && subtotal < 60 ? 8 : 0;
  var subtotalNode = document.getElementById('order-subtotal');
  var deliveryNode = document.getElementById('delivery-fee');
  var totalNode = document.getElementById('order-total');
  if (subtotalNode) subtotalNode.textContent = subtotal.toFixed(2) + ' ر.س';
  if (deliveryNode) deliveryNode.textContent = delivery ? delivery.toFixed(2) + ' ر.س' : 'مجانًا';
  if (totalNode) totalNode.textContent = (subtotal + delivery).toFixed(2) + ' ر.س';
}

document.addEventListener('DOMContentLoaded', function () {
  var toggle = document.querySelector('.menu-toggle');
  var nav = document.querySelector('.main-nav');
  if (toggle && nav) {
    toggle.addEventListener('click', function () {
      var isOpen = nav.classList.toggle('open');
      toggle.setAttribute('aria-expanded', String(isOpen));
    });
  }

  var registerForm = document.getElementById('register-form');
  if (registerForm) {
    registerForm.addEventListener('submit', function (event) {
      event.preventDefault();
      var name = registerForm.querySelector('[name="fullName"]').value.trim();
      var email = registerForm.querySelector('[name="email"]').value.trim();
      var password = registerForm.querySelector('[name="password"]').value;
      var goal = registerForm.querySelector('[name="goal"]').value;
      var terms = registerForm.querySelector('[name="terms"]').checked;
      if (!name || !email || !password || !goal || !terms || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        showStatus(registerForm, 'يرجى إدخال البيانات المطلوبة بصيغة صحيحة والموافقة على الإقرار.', 'error');
      } else if (password.length < 6) {
        showStatus(registerForm, 'كلمة المرور قصيرة؛ استخدمي 6 أحرف على الأقل.', 'error');
      } else {
        showStatus(registerForm, 'تم التحقق من بياناتك بنجاح يا ' + name + '.', 'success');
        window.alert('تم التحقق من بياناتك بنجاح. أهلًا بكِ يا ' + name + '!');
        registerForm.reset();
      }
    });
  }

  var orderForm = document.getElementById('order-form');
  if (orderForm) {
    var mealSelect = document.getElementById('order-meal');
    var quantityInput = document.getElementById('order-quantity');
    mealSelect.addEventListener('change', updateOrderSummary);
    quantityInput.addEventListener('input', updateOrderSummary);
    updateOrderSummary();
    orderForm.addEventListener('submit', function (event) {
      event.preventDefault();
      var name = orderForm.querySelector('[name="customerName"]').value.trim();
      var phone = orderForm.querySelector('[name="phone"]').value.trim();
      var address = orderForm.querySelector('[name="address"]').value.trim();
      var quantity = Number(quantityInput.value);
      var email = orderForm.querySelector('[name="email"]').value.trim();
      var validPhone = /^[0-9+\s-]{8,15}$/.test(phone);
      var validEmail = email === '' || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
      if (!name || !phone || !address) {
        showStatus(orderForm, 'أكملي الاسم ورقم الجوال والعنوان قبل إرسال الطلب.', 'error');
      } else if (!validPhone) {
        showStatus(orderForm, 'تحققي من رقم الجوال؛ يجب أن يتكون من 8 إلى 15 رقمًا.', 'error');
      } else if (!Number.isInteger(quantity) || quantity < 1 || quantity > 20) {
        showStatus(orderForm, 'اختاري كمية صحيحة من 1 إلى 20.', 'error');
      } else if (!validEmail) {
        showStatus(orderForm, 'تحققي من صيغة البريد الإلكتروني أو اتركيه فارغًا.', 'error');
      } else {
        showStatus(orderForm, 'تم إعداد ملخص الطلب يا ' + name + '. الإجمالي: ' + document.getElementById('order-total').textContent, 'success');
        showToast('شكرًا لكِ، طلبك قيد التجهيز!');
      }
    });
  }

  var calcRows = document.querySelectorAll('[data-calc-item] input');
  for (var j = 0; j < calcRows.length; j++) {
    calcRows[j].addEventListener('input', calculateOrderTotal);
  }
  if (calcRows.length) calculateOrderTotal();

  var year = document.getElementById('current-year');
  if (year) year.textContent = new Date().getFullYear();
});
