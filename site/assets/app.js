const products=[
{name:'Nissan Skyline GT-R (R34)',series:'Retro Racers / 2022',condition:'Mint',price:2950,stock:2,rarity:'Rare'},
{name:'Honda Civic EF',series:'Mainline / 2024',condition:'Good',price:499,stock:0,rarity:'Common'},
{name:'Porsche 911 GT3',series:'Premium / 2023',condition:'Mint',price:1499,stock:4,rarity:'Premium'},
{name:'Datsun 510 Wagon',series:'Treasure Hunt / 2021',condition:'Slightly damaged',price:2199,stock:1,rarity:'TH'}
];
function renderCards(id){const el=document.getElementById(id);if(!el)return;el.innerHTML=products.map(p=>`<article class='card'><div class='ph'></div><div class='content'><h3>${p.name}</h3><p class='meta'>${p.series}</p><p class='meta'>Condition: ${p.condition}</p><p class='price'>₹${p.price.toLocaleString('en-IN')}</p><p>${p.stock?`<span class='badge'>In stock (${p.stock})</span>`:`<span class='badge warn'>Out of stock</span>`}</p><div style='display:flex;gap:8px;flex-wrap:wrap'><a class='btn' href='product.html'>Add to Cart</a><a class='btn ghost' href='contact.html'>WhatsApp Enquiry</a></div></div></article>`).join('')}
document.addEventListener('DOMContentLoaded',()=>renderCards('productGrid'));
