let products = [];
let cart = [];
let currentCategory = 'all';
let currentProductId = null;

async function loadProducts() {
    try {
        const response = await fetch('products_clean.json');
        const data = await response.json();
        products = data.products;

        // Обновляем счетчик для кнопки "Все товары"
        const allBtn = document.querySelector('[data-category="all"] .category-count');
        if (allBtn) {
            allBtn.textContent = products.length;
        }

        renderCategories(data.categories);
        renderProducts(products);
        updateProductsCount(products.length);

        // Проверяем, есть ли товар в URL при загрузке страницы
        const productFromUrl = getProductFromUrl();
        if (productFromUrl) {
            showProductModal(productFromUrl);
        }

    } catch (error) {
        console.error('Error loading products:', error);
    }
}

function renderCategories(categories) {
    const wrapper = document.getElementById('categories-wrapper');

    // Подсчет товаров по категориям
    const categoryCounts = {};
    products.forEach(product => {
        if (!categoryCounts[product.category]) {
            categoryCounts[product.category] = 0;
        }
        categoryCounts[product.category]++;
    });

    categories.forEach(category => {
        const btn = document.createElement('button');
        btn.className = 'category-btn';
        btn.dataset.category = category;

        const count = categoryCounts[category] || 0;
        btn.innerHTML = `
            <span class="category-name">${category}</span>
            <span class="category-count">${count}</span>
        `;

        btn.addEventListener('click', () => filterByCategory(category));
        wrapper.appendChild(btn);
    });

    // Инициализация карусели
    initCategoriesCarousel();
}

function initCategoriesCarousel() {
    const wrapper = document.getElementById('categories-wrapper');
    const prevBtn = document.getElementById('scroll-prev');
    const nextBtn = document.getElementById('scroll-next');

    const scrollAmount = 200;

    function updateScrollButtons() {
        const isAtStart = wrapper.scrollLeft <= 0;
        const isAtEnd = wrapper.scrollLeft >= wrapper.scrollWidth - wrapper.clientWidth - 10;

        prevBtn.classList.toggle('hidden', isAtStart);
        nextBtn.classList.toggle('hidden', isAtEnd);
    }

    prevBtn.addEventListener('click', () => {
        wrapper.scrollBy({ left: -scrollAmount, behavior: 'smooth' });
        setTimeout(updateScrollButtons, 300);
    });

    nextBtn.addEventListener('click', () => {
        wrapper.scrollBy({ left: scrollAmount, behavior: 'smooth' });
        setTimeout(updateScrollButtons, 300);
    });

    wrapper.addEventListener('scroll', updateScrollButtons);

    // Проверка при загрузке
    setTimeout(updateScrollButtons, 100);

    // Проверка при изменении размера окна
    window.addEventListener('resize', updateScrollButtons);
}

function filterByCategory(category) {
    currentCategory = category;

    document.querySelectorAll('.category-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.category === category);
    });

    const filtered = category === 'all'
        ? products
        : products.filter(p => p.category === category);

    renderProducts(filtered);
    updateProductsCount(filtered.length);
    document.getElementById('category-title').textContent =
        category === 'all' ? 'Все товары' : category;
}

function renderProducts(productList) {
    const container = document.getElementById('products-container');
    container.innerHTML = '';

    productList.forEach(product => {
        const card = createProductCard(product);
        container.appendChild(card);
    });
}

function createProductCard(product) {
    const card = document.createElement('div');
    card.className = 'product-card';
    card.style.opacity = '0';
    card.style.transform = 'translateY(20px)';
    card.onclick = () => showProductModal(product);

    card.innerHTML = `
        <div class="product-image">
            <img src="${product.image}" alt="${product.name}"
                 onerror="this.src='images/no-image.png'">
        </div>
        <div class="product-info">
            <h3 class="product-name">${product.name}</h3>
            <p class="product-article">Артикул: ${product.article || 'Не указан'}</p>
            <p class="product-price">${formatPrice(product.price)}</p>
            <span class="product-stock ${product.in_stock ? 'in-stock' : 'out-of-stock'}">
                ${product.in_stock ? `В наличии (${product.quantity} шт.)` : 'Под заказ'}
            </span>
        </div>
    `;

    // Анимация появления
    setTimeout(() => {
        card.style.transition = 'all 0.4s ease-out';
        card.style.opacity = '1';
        card.style.transform = 'translateY(0)';
    }, 50);

    return card;
}

function showProductModal(product) {
    const modal = document.getElementById('product-modal');
    currentProductId = product.id;

    // Создаем URL-friendly slug из названия товара
    const slug = createSlug(product.name);
    const productUrl = `#product-${product.id}-${slug}`;

    // Обновляем URL без перезагрузки страницы
    history.pushState(
        { productId: product.id, type: 'product' },
        `${product.name} - СанТехКаталог`,
        productUrl
    );

    modal.querySelector('.modal-image img').src = product.image;
    modal.querySelector('.modal-title').textContent = product.name;
    modal.querySelector('.modal-article').textContent = `Артикул: ${product.article || 'Не указан'}`;
    modal.querySelector('.modal-price').textContent = formatPrice(product.price);
    modal.querySelector('.modal-stock').innerHTML = `
        <span class="product-stock ${product.in_stock ? 'in-stock' : 'out-of-stock'}">
            ${product.in_stock ? `В наличии: ${product.quantity} шт.` : 'Под заказ'}
        </span>
    `;
    modal.querySelector('.modal-description').textContent = product.description || 'Описание отсутствует';

    const addBtn = modal.querySelector('.add-to-cart-btn');
    addBtn.onclick = () => {
        addToCart(product);
        closeModal('product-modal');
    };

    modal.classList.add('active');

    // Обновляем заголовок страницы
    document.title = `${product.name} - AquaTek - Профессиональная сантехника`;
}

function addToCart(product) {
    const existingItem = cart.find(item => item.id === product.id);

    if (existingItem) {
        existingItem.quantity++;
    } else {
        cart.push({
            ...product,
            quantity: 1
        });
    }

    updateCartCount();
    showNotification('Товар добавлен в корзину');
}

function updateCartCount() {
    const count = cart.reduce((sum, item) => sum + item.quantity, 0);
    document.querySelector('.cart-count').textContent = count;
}

function showCart() {
    const modal = document.getElementById('cart-modal');
    const itemsContainer = document.getElementById('cart-items');

    if (cart.length === 0) {
        itemsContainer.innerHTML = '<div class="empty-cart">Корзина пуста</div>';
    } else {
        itemsContainer.innerHTML = cart.map(item => `
            <div class="cart-item">
                <div class="cart-item-image">
                    <img src="${item.image}" alt="${item.name}">
                </div>
                <div class="cart-item-info">
                    <div class="cart-item-name">${item.name}</div>
                    <div class="cart-item-price">${formatPrice(item.price)}</div>
                    <div class="cart-item-quantity">
                        <button class="quantity-btn" onclick="updateQuantity(${item.id}, -1)">-</button>
                        <span>${item.quantity}</span>
                        <button class="quantity-btn" onclick="updateQuantity(${item.id}, 1)">+</button>
                    </div>
                </div>
                <button class="cart-item-remove" onclick="removeFromCart(${item.id})">Удалить</button>
            </div>
        `).join('');
    }

    updateCartTotal();
    modal.classList.add('active');
}

function updateQuantity(productId, change) {
    const item = cart.find(i => i.id === productId);
    if (item) {
        item.quantity += change;
        if (item.quantity <= 0) {
            removeFromCart(productId);
        } else {
            showCart();
            updateCartCount();
        }
    }
}

function removeFromCart(productId) {
    cart = cart.filter(item => item.id !== productId);
    showCart();
    updateCartCount();
}

function updateCartTotal() {
    const total = cart.reduce((sum, item) => sum + (item.price * item.quantity), 0);
    document.getElementById('cart-total').textContent = formatPrice(total);
}

function closeModal(modalId) {
    document.getElementById(modalId).classList.remove('active');

    if (modalId === 'product-modal' && currentProductId) {
        // Возвращаемся к главной странице
        history.pushState(
            { type: 'main' },
            'AquaTek - Профессиональная сантехника',
            window.location.pathname
        );
        document.title = 'AquaTek - Профессиональная сантехника';
        currentProductId = null;
    }
}

function formatPrice(price) {
    return new Intl.NumberFormat('ru-RU', {
        style: 'currency',
        currency: 'RUB',
        minimumFractionDigits: 0,
        maximumFractionDigits: 2
    }).format(price);
}

function updateProductsCount(count) {
    document.querySelector('.products-count').textContent = `Найдено: ${count}`;
}

function createSlug(text) {
    return text
        .toLowerCase()
        .replace(/[^\w\s-]/g, '') // удаляем специальные символы
        .replace(/\s+/g, '-') // заменяем пробелы на дефисы
        .replace(/-+/g, '-') // заменяем множественные дефисы на один
        .trim();
}

function getProductFromUrl() {
    const hash = window.location.hash;
    if (hash.startsWith('#product-')) {
        const productId = parseInt(hash.split('-')[1]);
        return products.find(p => p.id === productId);
    }
    return null;
}

function handleBrowserNavigation(event) {
    if (event.state) {
        if (event.state.type === 'product' && event.state.productId) {
            const product = products.find(p => p.id === event.state.productId);
            if (product) {
                showProductModal(product);
            }
        } else if (event.state.type === 'main') {
            closeModal('product-modal');
        }
    } else {
        // Пользователь нажал назад на главной странице
        closeModal('product-modal');
    }
}

function shareProduct() {
    const currentUrl = window.location.href;

    if (navigator.share) {
        // Используем нативный API для мобильных устройств
        navigator.share({
            title: document.title,
            text: 'Посмотрите этот товар в AquaTek - профессиональная сантехника',
            url: currentUrl
        });
    } else {
        // Копируем ссылку в буфер обмена для десктопа
        navigator.clipboard.writeText(currentUrl).then(() => {
            showNotification('Ссылка скопирована в буфер обмена!');
        }).catch(() => {
            // Fallback для старых браузеров
            const textArea = document.createElement('textarea');
            textArea.value = currentUrl;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            showNotification('Ссылка скопирована в буфер обмена!');
        });
    }
}

function showNotification(message) {
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: #10b981;
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        z-index: 2000;
        animation: slideIn 0.3s ease-out;
    `;
    notification.textContent = message;
    document.body.appendChild(notification);

    setTimeout(() => {
        notification.remove();
    }, 3000);
}

document.addEventListener('DOMContentLoaded', function() {
    loadProducts();

    // Добавляем обработчик для кнопок "назад/вперед" браузера
    window.addEventListener('popstate', handleBrowserNavigation);

    document.querySelector('.cart-button').addEventListener('click', showCart);

    document.querySelectorAll('.modal-close').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const modal = e.target.closest('.product-modal, .cart-modal');
            if (modal) {
                if (modal.classList.contains('product-modal')) {
                    closeModal('product-modal');
                } else {
                    closeModal('cart-modal');
                }
            }
        });
    });

    document.getElementById('search').addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const filtered = products.filter(p =>
            p.name.toLowerCase().includes(query) ||
            (p.article && p.article.toLowerCase().includes(query))
        );
        renderProducts(filtered);
        updateProductsCount(filtered.length);
    });

    document.getElementById('in-stock').addEventListener('change', (e) => {
        let filtered = currentCategory === 'all'
            ? products
            : products.filter(p => p.category === currentCategory);

        if (e.target.checked) {
            filtered = filtered.filter(p => p.in_stock);
        }

        renderProducts(filtered);
        updateProductsCount(filtered.length);
    });

    document.getElementById('sort-select').addEventListener('change', (e) => {
        const container = document.getElementById('products-container');
        const cards = Array.from(container.children);
        const sortedProducts = [...products];

        switch(e.target.value) {
            case 'price-asc':
                sortedProducts.sort((a, b) => a.price - b.price);
                break;
            case 'price-desc':
                sortedProducts.sort((a, b) => b.price - a.price);
                break;
            case 'name':
                sortedProducts.sort((a, b) => a.name.localeCompare(b.name));
                break;
        }

        renderProducts(sortedProducts);
    });

    document.getElementById('min-price').addEventListener('input', filterByPrice);
    document.getElementById('max-price').addEventListener('input', filterByPrice);

    document.querySelector('.reset-filters').addEventListener('click', () => {
        document.getElementById('search').value = '';
        document.getElementById('in-stock').checked = false;
        document.getElementById('min-price').value = '';
        document.getElementById('max-price').value = '';
        document.getElementById('sort-select').value = 'name';

        filterByCategory('all');
    });

    document.querySelector('.checkout-btn').addEventListener('click', () => {
        if (cart.length > 0) {
            alert('Спасибо за заказ! Мы свяжемся с вами в ближайшее время.');
            cart = [];
            updateCartCount();
            closeModal('cart-modal');
        }
    });

    document.querySelectorAll('.category-btn')[0].classList.add('active');
});

function filterByPrice() {
    const min = parseFloat(document.getElementById('min-price').value) || 0;
    const max = parseFloat(document.getElementById('max-price').value) || Infinity;

    let filtered = currentCategory === 'all'
        ? products
        : products.filter(p => p.category === currentCategory);

    filtered = filtered.filter(p => p.price >= min && p.price <= max);

    if (document.getElementById('in-stock').checked) {
        filtered = filtered.filter(p => p.in_stock);
    }

    renderProducts(filtered);
    updateProductsCount(filtered.length);
}

const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
`;
document.head.appendChild(style);