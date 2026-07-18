import React, { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Plus, Minus, Trash2, Download, ClipboardList, UtensilsCrossed, X, ChefHat, Table2, ShoppingBag, RotateCcw, Receipt, Clock, ArrowRight, Printer } from 'lucide-react';
import { supabase, supabaseConfigured } from './supabaseClient';

const CATEGORIES = ['Hamburguesas', 'Salchipapa', 'Salchibroaster', 'Salchialitas', 'Broaster', 'Mostrito', 'Sandwich Broaster', 'Alitas', 'Enchiladas', 'Jugos', 'Bebidas', 'Calientes', 'Extras'];

const DEFAULT_MENU = [
  { name: 'Clásica (Carne)', category: 'Hamburguesas', price: 10.0 },
  { name: 'Clásica (Pollo/Chorizo)', category: 'Hamburguesas', price: 10.5 },
  { name: 'Especial (Carne)', category: 'Hamburguesas', price: 11.5 },
  { name: 'Especial (Pollo/Chorizo)', category: 'Hamburguesas', price: 12.0 },
  { name: 'Americana (Carne)', category: 'Hamburguesas', price: 11.0 },
  { name: 'Americana (Pollo/Chorizo)', category: 'Hamburguesas', price: 11.5 },
  { name: 'Royal (Carne)', category: 'Hamburguesas', price: 12.5 },
  { name: 'Royal (Pollo/Chorizo)', category: 'Hamburguesas', price: 13.0 },
  { name: 'Mixta (Carne)', category: 'Hamburguesas', price: 12.5 },
  { name: 'Mixta (Pollo/Chorizo)', category: 'Hamburguesas', price: 13.0 },
  { name: 'A lo Pobre (Carne)', category: 'Hamburguesas', price: 12.0 },
  { name: 'A lo Pobre (Pollo/Chorizo)', category: 'Hamburguesas', price: 12.5 },
  { name: 'Royal Mixta (Carne)', category: 'Hamburguesas', price: 13.0 },
  { name: 'Royal Mixta (Pollo/Chorizo)', category: 'Hamburguesas', price: 13.5 },
  { name: 'Mixta Especial (Carne)', category: 'Hamburguesas', price: 13.5 },
  { name: 'Mixta Especial (Pollo/Chorizo)', category: 'Hamburguesas', price: 14.0 },
  { name: 'Hawaiiana (Carne)', category: 'Hamburguesas', price: 14.0 },
  { name: 'Hawaiiana (Pollo/Chorizo)', category: 'Hamburguesas', price: 14.5 },
  { name: 'Royal Especial (Carne)', category: 'Hamburguesas', price: 14.5 },
  { name: 'Royal Especial (Pollo/Chorizo)', category: 'Hamburguesas', price: 15.0 },
  { name: 'Clásica', category: 'Salchipapa', price: 11.0 },
  { name: 'c/ Huevo', category: 'Salchipapa', price: 12.0 },
  { name: 'c/ Queso', category: 'Salchipapa', price: 13.0 },
  { name: 'c/ Chorizo', category: 'Salchipapa', price: 13.0 },
  { name: 'c/ Pollo Desilachado', category: 'Salchipapa', price: 14.0 },
  { name: 'c/ Tocino', category: 'Salchipapa', price: 13.5 },
  { name: 'c/ Cabanossi', category: 'Salchipapa', price: 14.0 },
  { name: 'c/ Cabanossi y Huevo', category: 'Salchipapa', price: 14.5 },
  { name: 'c/ Nuggets (4 und)', category: 'Salchipapa', price: 15.0 },
  { name: 'Mixta (huevo y tocino)', category: 'Salchipapa', price: 14.5 },
  { name: 'Especial (chorizo, huevo y tocino)', category: 'Salchipapa', price: 16.5 },
  { name: 'A lo Pobre (plátano y huevo)', category: 'Salchipapa', price: 13.5 },
  { name: 'c/ Chicharrón de Pollo', category: 'Salchipapa', price: 15.5 },
  { name: 'Ala', category: 'Salchibroaster', price: 15.0 },
  { name: 'Pierna', category: 'Salchibroaster', price: 16.0 },
  { name: 'Encuentro', category: 'Salchibroaster', price: 16.5 },
  { name: 'Pecho', category: 'Salchibroaster', price: 17.5 },
  { name: '4 Alitas + Salchipapa + Ensalada (1 sabor)', category: 'Salchialitas', price: 17.5 },
  { name: '6 Alitas + Salchipapa + Ensalada (2 sabores)', category: 'Salchialitas', price: 23.5 },
  { name: 'Ala', category: 'Broaster', price: 12.5 },
  { name: 'Pierna', category: 'Broaster', price: 13.5 },
  { name: 'Encuentro', category: 'Broaster', price: 14.0 },
  { name: 'Pecho', category: 'Broaster', price: 15.0 },
  { name: 'Ala', category: 'Mostrito', price: 16.0 },
  { name: 'Pierna', category: 'Mostrito', price: 17.0 },
  { name: 'Encuentro', category: 'Mostrito', price: 17.5 },
  { name: 'Pecho', category: 'Mostrito', price: 18.5 },
  { name: 'Sandwich Broaster Clásico', category: 'Sandwich Broaster', price: 15.0 },
  { name: 'Sandwich Broaster c/ Salsa Especial', category: 'Sandwich Broaster', price: 16.5 },
  { name: '6 Alitas + Papas + Ensalada', category: 'Alitas', price: 17.0 },
  { name: '18 Alitas + Papas + Ensalada (3 sabores)', category: 'Alitas', price: 49.5 },
  { name: '24 Alitas + Papas + Ensalada (4 sabores)', category: 'Alitas', price: 63.0 },
  { name: '20 Alitas (sin papas ni ensalada)', category: 'Alitas', price: 49.9 },
  { name: '50 Alitas (sin papas ni ensalada)', category: 'Alitas', price: 105.9 },
  { name: '6 Alitas + Chaufa + Papas + Ensalada', category: 'Alitas', price: 20.0 },
  { name: 'Enchilada de Pollo', category: 'Enchiladas', price: 13.0 },
  { name: 'Enchilada de Chorizo', category: 'Enchiladas', price: 13.0 },
  { name: 'Enchilada de Carne', category: 'Enchiladas', price: 13.0 },
  { name: 'Enchilada Especial', category: 'Enchiladas', price: 14.5 },
  { name: 'Enchilada Mixta', category: 'Enchiladas', price: 15.0 },
  { name: 'Enchilada Completa', category: 'Enchiladas', price: 18.5 },
  { name: 'Enchilada Hawaiana', category: 'Enchiladas', price: 16.0 },
  { name: 'Twister', category: 'Enchiladas', price: 15.0 },
  { name: 'Papaya', category: 'Jugos', price: 8.0 },
  { name: 'Surtido', category: 'Jugos', price: 8.0 },
  { name: 'Piña', category: 'Jugos', price: 8.0 },
  { name: 'Especial', category: 'Jugos', price: 10.0 },
  { name: 'Fresa', category: 'Jugos', price: 9.0 },
  { name: 'Plátano c/ leche', category: 'Jugos', price: 10.0 },
  { name: 'Mango', category: 'Jugos', price: 10.0 },
  { name: 'Lúcuma c/ leche', category: 'Jugos', price: 12.0 },
  { name: 'Fresa c/ leche', category: 'Jugos', price: 10.0 },
  { name: 'Mango c/ leche', category: 'Jugos', price: 12.0 },
  { name: 'Papaya c/ leche', category: 'Jugos', price: 10.0 },
  { name: 'Naranja', category: 'Jugos', price: 7.0 },
  { name: 'Chicha Morada (vaso)', category: 'Bebidas', price: 7.0 },
  { name: 'Chicha Morada (jarra)', category: 'Bebidas', price: 12.0 },
  { name: 'Limonada (vaso)', category: 'Bebidas', price: 8.0 },
  { name: 'Limonada (jarra)', category: 'Bebidas', price: 12.0 },
  { name: 'Limonada Frozen (vaso)', category: 'Bebidas', price: 10.0 },
  { name: 'Limonada Frozen (jarra)', category: 'Bebidas', price: 16.0 },
  { name: 'Maracuyá (vaso)', category: 'Bebidas', price: 8.0 },
  { name: 'Maracuyá (jarra)', category: 'Bebidas', price: 12.0 },
  { name: 'Maracuyá Frozen (vaso)', category: 'Bebidas', price: 10.0 },
  { name: 'Maracuyá Frozen (jarra)', category: 'Bebidas', price: 16.0 },
  { name: 'Inca/Coca Personal', category: 'Bebidas', price: 2.0 },
  { name: 'Inca/Coca Litro', category: 'Bebidas', price: 6.0 },
  { name: 'Gordita', category: 'Bebidas', price: 4.5 },
  { name: 'Pepsi ½ L', category: 'Bebidas', price: 2.5 },
  { name: 'Pepsi 1 L', category: 'Bebidas', price: 4.5 },
  { name: 'Té / Anís / Manzanilla', category: 'Calientes', price: 2.5 },
  { name: 'Café', category: 'Calientes', price: 3.5 },
  { name: 'Café c/ leche', category: 'Calientes', price: 5.0 },
  { name: 'Porción de Arroz', category: 'Extras', price: 2.0 },
];

const todayISO = () => new Date().toISOString().slice(0, 10);
const weekAgoISO = () => { const d = new Date(); d.setDate(d.getDate() - 6); return d.toISOString().slice(0, 10); };
const fmt = (n) => `S/ ${Number(n).toFixed(2)}`;
const zigzag = 'polygon(0% 0%,100% 0%,100% 90%,95% 100%,90% 90%,85% 100%,80% 90%,75% 100%,70% 90%,65% 100%,60% 90%,55% 100%,50% 90%,45% 100%,40% 90%,35% 100%,30% 90%,25% 100%,20% 90%,15% 100%,10% 90%,5% 100%,0% 90%)';

function mapSaleRow(row) {
  return {
    id: row.id,
    date: row.sale_date,
    time: row.sale_time,
    type: row.order_type,
    table: row.table_number,
    items: row.items,
    total: Number(row.total),
    payment: row.payment_method || 'efectivo',
  };
}

export default function App() {
  if (!supabaseConfigured) {
    return (
      <div style={{ fontFamily: 'sans-serif', padding: 40, maxWidth: 600, margin: '0 auto', color: '#333' }}>
        <h2>Falta configurar Supabase</h2>
        <p>Esta app necesita las variables de entorno <code>VITE_SUPABASE_URL</code> y <code>VITE_SUPABASE_ANON_KEY</code>.</p>
        <p>Revisa el archivo <code>README.md</code> del proyecto para los pasos de configuración en Vercel.</p>
      </div>
    );
  }

  const [tab, setTab] = useState('pedido');
  const [menu, setMenu] = useState([]);
  const [orders, setOrders] = useState([]);
  const [openOrdersData, setOpenOrdersData] = useState({}); // { key: { items, notes, opened_at } }
  const [numTables, setNumTables] = useState(12);
  const [loaded, setLoaded] = useState(false);

  const [selectedKey, setSelectedKey] = useState('llevar');
  const [activeCat, setActiveCat] = useState(CATEGORIES[0]);
  const [confirmFlash, setConfirmFlash] = useState(false);
  const [printOrder, setPrintOrder] = useState(null);

  const [dateFrom, setDateFrom] = useState(todayISO());
  const [dateTo, setDateTo] = useState(todayISO());

  const [newProd, setNewProd] = useState({ name: '', category: CATEGORIES[0], price: '' });
  const [menuCat, setMenuCat] = useState('Todos');

  // Initial load
  useEffect(() => {
    (async () => {
      const { data: menuData } = await supabase.from('menu_items').select('*').order('category').order('name');
      if (!menuData || menuData.length === 0) {
        await supabase.from('menu_items').insert(DEFAULT_MENU);
        const { data: seeded } = await supabase.from('menu_items').select('*').order('category').order('name');
        setMenu(seeded || []);
      } else {
        setMenu(menuData);
      }

      const { data: openData } = await supabase.from('open_orders').select('*');
      const openMap = {};
      (openData || []).forEach((r) => { openMap[r.order_key] = { items: r.items || [], notes: r.notes || '', opened_at: r.opened_at }; });
      setOpenOrdersData(openMap);

      const { data: salesData } = await supabase.from('sales').select('*').order('created_at', { ascending: false }).limit(3000);
      setOrders((salesData || []).map(mapSaleRow));

      const { data: settingsData } = await supabase.from('app_settings').select('*').eq('key', 'num_tables').maybeSingle();
      if (settingsData && settingsData.value && settingsData.value.n) setNumTables(settingsData.value.n);
      else await supabase.from('app_settings').insert({ key: 'num_tables', value: { n: 12 } });

      setLoaded(true);
    })();
  }, []);

  // Realtime subscriptions
  useEffect(() => {
    const channel = supabase
      .channel('realtime-compachos')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'open_orders' }, (payload) => {
        if (payload.eventType === 'DELETE') {
          setOpenOrdersData((prev) => { const n = { ...prev }; delete n[payload.old.order_key]; return n; });
        } else {
          const row = payload.new;
          setOpenOrdersData((prev) => ({ ...prev, [row.order_key]: { items: row.items || [], notes: row.notes || '', opened_at: row.opened_at } }));
        }
      })
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'sales' }, (payload) => {
        setOrders((prev) => (prev.some((o) => o.id === payload.new.id) ? prev : [mapSaleRow(payload.new), ...prev]));
      })
      .on('postgres_changes', { event: 'DELETE', schema: 'public', table: 'sales' }, (payload) => {
        setOrders((prev) => prev.filter((o) => o.id !== payload.old.id));
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_items' }, (payload) => {
        if (payload.eventType === 'DELETE') {
          setMenu((prev) => prev.filter((m) => m.id !== payload.old.id));
        } else {
          setMenu((prev) => {
            const exists = prev.find((m) => m.id === payload.new.id);
            if (exists) return prev.map((m) => (m.id === payload.new.id ? payload.new : m));
            return [...prev, payload.new];
          });
        }
      })
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, []);

  useEffect(() => {
    function handleAfterPrint() { setPrintOrder(null); }
    window.addEventListener('afterprint', handleAfterPrint);
    return () => window.removeEventListener('afterprint', handleAfterPrint);
  }, []);

  useEffect(() => {
    if (printOrder) {
      const id = setTimeout(() => window.print(), 80);
      return () => clearTimeout(id);
    }
  }, [printOrder]);

  const cartData = openOrdersData[selectedKey] || { items: [], notes: '' };
  const cart = cartData.items;
  const cartTotal = useMemo(() => cart.reduce((s, l) => s + l.price * l.qty, 0), [cart]);
  const isTable = selectedKey.startsWith('mesa-');
  const tableNum = isTable ? selectedKey.replace('mesa-', '') : null;

  const saveOpenOrder = useCallback(async (key, items, notes) => {
    if (items.length === 0) {
      await supabase.from('open_orders').delete().eq('order_key', key);
      setOpenOrdersData((prev) => { const n = { ...prev }; delete n[key]; return n; });
      return;
    }
    const opened_at = openOrdersData[key]?.opened_at || new Date().toISOString();
    await supabase.from('open_orders').upsert({ order_key: key, items, notes: notes || '', opened_at }, { onConflict: 'order_key' });
    setOpenOrdersData((prev) => ({ ...prev, [key]: { items, notes: notes || '', opened_at } }));
  }, [openOrdersData]);

  function addToCart(item) {
    const current = cart;
    const found = current.find((l) => l.menuId === item.id);
    const next = found
      ? current.map((l) => (l.menuId === item.id ? { ...l, qty: l.qty + 1 } : l))
      : [...current, { menuId: item.id, name: item.name, category: item.category, price: Number(item.price), qty: 1 }];
    saveOpenOrder(selectedKey, next, cartData.notes);
  }
  function changeQty(menuId, delta) {
    const next = cart.map((l) => (l.menuId === menuId ? { ...l, qty: l.qty + delta } : l)).filter((l) => l.qty > 0);
    saveOpenOrder(selectedKey, next, cartData.notes);
  }
  function removeLine(menuId) {
    saveOpenOrder(selectedKey, cart.filter((l) => l.menuId !== menuId), cartData.notes);
  }
  function updateNotes(text) {
    setOpenOrdersData((prev) => ({ ...prev, [selectedKey]: { ...(prev[selectedKey] || { items: [], opened_at: new Date().toISOString() }), notes: text } }));
  }
  function commitNotes() {
    if (cart.length > 0) saveOpenOrder(selectedKey, cart, cartData.notes);
  }

  async function chargeOrderForKey(key, paymentMethod) {
    const data = openOrdersData[key];
    if (!data || data.items.length === 0) return;
    const total = data.items.reduce((s, l) => s + l.price * l.qty, 0);
    const now = new Date();
    const keyIsTable = key.startsWith('mesa-');
    const sale = {
      sale_date: now.toISOString().slice(0, 10),
      sale_time: now.toTimeString().slice(0, 5),
      order_type: keyIsTable ? 'mesa' : 'llevar',
      table_number: keyIsTable ? key.replace('mesa-', '') : null,
      items: data.items,
      total,
      payment_method: paymentMethod,
    };
    const { data: inserted, error } = await supabase.from('sales').insert(sale).select();
    if (error) {
      console.error('Error al guardar la venta:', error);
      alert('No se pudo guardar la venta en el Historial. La mesa/pedido NO se liberó, así que no perdiste nada — inténtalo de nuevo.\n\nDetalle técnico: ' + error.message);
      return;
    }
    if (inserted && inserted[0]) setOrders((prev) => [mapSaleRow(inserted[0]), ...prev]);
    await supabase.from('open_orders').delete().eq('order_key', key);
    setOpenOrdersData((prev) => { const n = { ...prev }; delete n[key]; return n; });
    setConfirmFlash(true);
    setTimeout(() => setConfirmFlash(false), 1400);
  }
  function chargeOrder(paymentMethod) { chargeOrderForKey(selectedKey, paymentMethod); }

  async function clearCartForKey(key) {
    const data = openOrdersData[key];
    if (!data || data.items.length === 0) return;
    if (!window.confirm('¿Vaciar este pedido sin cobrarlo?')) return;
    await supabase.from('open_orders').delete().eq('order_key', key);
    setOpenOrdersData((prev) => { const n = { ...prev }; delete n[key]; return n; });
  }
  function clearCart() { clearCartForKey(selectedKey); }

  function printComanda(key) {
    const data = openOrdersData[key];
    if (!data || data.items.length === 0) return;
    const now = new Date();
    setPrintOrder({
      label: key === 'llevar' ? 'PARA LLEVAR' : `MESA ${key.replace('mesa-', '')}`,
      time: now.toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' }),
      date: now.toLocaleDateString('es-PE'),
      items: data.items,
      notes: data.notes || '',
    });
  }

  function minutesAgo(iso) {
    if (!iso) return 0;
    return Math.max(0, Math.round((Date.now() - new Date(iso).getTime()) / 60000));
  }

  const openKeys = useMemo(() => {
    return Object.keys(openOrdersData)
      .filter((k) => (openOrdersData[k]?.items || []).length > 0)
      .sort((a, b) => {
        if (a === 'llevar') return -1;
        if (b === 'llevar') return 1;
        return Number(a.replace('mesa-', '')) - Number(b.replace('mesa-', ''));
      });
  }, [openOrdersData]);

  const filteredOrders = useMemo(
    () => orders.filter((o) => o.date >= dateFrom && o.date <= dateTo).sort((a, b) => (a.date + a.time < b.date + b.time ? 1 : -1)),
    [orders, dateFrom, dateTo]
  );
  const totalVentas = filteredOrders.reduce((s, o) => s + o.total, 0);
  const ventasLlevar = filteredOrders.filter((o) => o.type === 'llevar').reduce((s, o) => s + o.total, 0);
  const ventasMesa = filteredOrders.filter((o) => o.type === 'mesa').reduce((s, o) => s + o.total, 0);
  const ventasEfectivo = filteredOrders.filter((o) => o.payment !== 'yape').reduce((s, o) => s + o.total, 0);
  const ventasYape = filteredOrders.filter((o) => o.payment === 'yape').reduce((s, o) => s + o.total, 0);
  const ticketProm = filteredOrders.length ? totalVentas / filteredOrders.length : 0;

  function exportExcel() {
    const byDay = {};
    filteredOrders.forEach((o) => {
      byDay[o.date] = byDay[o.date] || { fecha: o.date, pedidos: 0, llevar: 0, mesa: 0, efectivo: 0, yape: 0, total: 0 };
      byDay[o.date].pedidos += 1;
      byDay[o.date][o.type === 'llevar' ? 'llevar' : 'mesa'] += o.total;
      byDay[o.date][o.payment === 'yape' ? 'yape' : 'efectivo'] += o.total;
      byDay[o.date].total += o.total;
    });
    const resumen = Object.values(byDay).sort((a, b) => (a.fecha < b.fecha ? 1 : -1)).map((d) => ({
      Fecha: d.fecha, 'N° Pedidos': d.pedidos, 'Ventas Para Llevar (S/)': d.llevar.toFixed(2),
      'Ventas Mesa (S/)': d.mesa.toFixed(2), 'Efectivo (S/)': d.efectivo.toFixed(2), 'Yape (S/)': d.yape.toFixed(2),
      'Total del Día (S/)': d.total.toFixed(2),
    }));
    const detalle = filteredOrders.map((o) => ({
      Fecha: o.date, Hora: o.time, Tipo: o.type === 'llevar' ? 'Para llevar' : `Mesa ${o.table}`,
      Pago: o.payment === 'yape' ? 'Yape' : 'Efectivo',
      Productos: o.items.map((i) => `${i.qty}x ${i.name}`).join('; '),
      'Total (S/)': o.total.toFixed(2),
    }));
    const prodMap = {};
    filteredOrders.forEach((o) => o.items.forEach((i) => {
      prodMap[i.name] = prodMap[i.name] || { producto: i.name, categoria: i.category, cantidad: 0, ingresos: 0 };
      prodMap[i.name].cantidad += i.qty;
      prodMap[i.name].ingresos += i.qty * i.price;
    }));
    const productos = Object.values(prodMap).sort((a, b) => b.ingresos - a.ingresos).map((p) => ({
      Producto: p.producto, Categoría: p.categoria, 'Cantidad Vendida': p.cantidad, 'Ingresos (S/)': p.ingresos.toFixed(2),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), 'Resumen por día');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detalle), 'Detalle de pedidos');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(productos), 'Productos vendidos');
    XLSX.writeFile(wb, `Ventas_${dateFrom}_a_${dateTo}.xlsx`);
  }

  async function deleteSale(id) {
    await supabase.from('sales').delete().eq('id', id);
    setOrders((prev) => prev.filter((o) => o.id !== id));
  }

  async function updateMenuItem(id, field, value) {
    const val = field === 'price' ? Number(value) : value;
    setMenu((prev) => prev.map((it) => (it.id === id ? { ...it, [field]: val } : it)));
    await supabase.from('menu_items').update({ [field]: val }).eq('id', id);
  }
  async function deleteMenuItem(id) {
    setMenu((prev) => prev.filter((it) => it.id !== id));
    await supabase.from('menu_items').delete().eq('id', id);
  }
  async function addProduct() {
    if (!newProd.name.trim() || !newProd.price) return;
    const { data } = await supabase.from('menu_items').insert({ name: newProd.name.trim(), category: newProd.category, price: Number(newProd.price) }).select();
    if (data && data[0]) setMenu((prev) => [...prev, data[0]]);
    setNewProd({ name: '', category: newProd.category, price: '' });
  }

  async function changeNumTables(n) {
    const val = Math.max(1, n);
    setNumTables(val);
    await supabase.from('app_settings').upsert({ key: 'num_tables', value: { n: val } });
  }

  const menuByCat = menuCat === 'Todos' ? menu : menu.filter((m) => m.category === menuCat);

  if (!loaded) {
    return (
      <div style={{ fontFamily: 'sans-serif', minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#20262a', color: '#eee9de' }}>
        Cargando datos…
      </div>
    );
  }

  return (
    <div className="app-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

        * { box-sizing:border-box; }
        body { margin:0; }
        .app-root { --steel-900:#20262a; --steel-800:#2b3237; --steel-700:#3a444a; --steel-600:#4a565d;
          --paper:#f6f1e2; --paper-line:#d9cfb4; --mustard:#e7a63c; --mustard-dark:#c98a24;
          --chili:#d8432e; --chili-dark:#b8341f; --ink:#241f18; --cream:#eee9de; --muted:#9aa7ac; --green:#4f9a63;
          font-family:'Inter',sans-serif; background:var(--steel-900); color:var(--cream); min-height:100vh;
        }
        .display { font-family:'Bebas Neue',sans-serif; letter-spacing:0.02em; }
        .mono { font-family:'IBM Plex Mono',monospace; }

        .topbar { background:var(--steel-800); border-bottom:1px solid var(--steel-700); padding:16px 22px;
          display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:14px; }
        .brand { display:flex; align-items:center; gap:10px; }
        .brand h1 { font-size:30px; margin:0; color:var(--mustard); line-height:1; }
        .brand .tag { font-size:11px; color:var(--muted); font-style:italic; }
        .tabs { display:flex; gap:6px; background:var(--steel-900); padding:5px; border-radius:12px; }
        .tab-btn { background:transparent; border:none; color:var(--muted); padding:12px 18px; border-radius:9px;
          font-weight:700; font-size:14px; cursor:pointer; display:flex; align-items:center; gap:7px; min-height:44px; }
        .tab-btn.active { background:var(--mustard); color:var(--ink); }
        .badge-count { background:var(--chili); color:#fff; font-size:11px; font-weight:700; border-radius:10px;
          padding:2px 7px; margin-left:2px; }

        .content { padding:22px; max-width:1400px; margin:0 auto; }
        .section-label { font-size:12px; text-transform:uppercase; letter-spacing:0.06em; color:var(--muted); font-weight:700; margin-bottom:10px; }

        .table-picker { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:8px; }
        .table-grid { display:flex; gap:10px; flex-wrap:wrap; }
        .table-btn { min-width:76px; min-height:64px; border-radius:12px; border:2px solid var(--steel-700); background:var(--steel-800);
          color:var(--cream); font-weight:700; font-size:15px; cursor:pointer; display:flex; flex-direction:column; align-items:center;
          justify-content:center; gap:2px; padding:8px; position:relative; }
        .table-btn.occupied { border-color:var(--mustard); background:#3a2f1c; }
        .table-btn.selected { border-color:var(--chili); box-shadow:0 0 0 2px var(--chili); }
        .table-btn .amt { font-family:'IBM Plex Mono',monospace; font-size:12px; color:var(--mustard); }
        .llevar-btn { min-width:120px; min-height:64px; border-radius:12px; border:2px solid var(--steel-700); background:var(--steel-800);
          color:var(--cream); font-weight:700; font-size:15px; cursor:pointer; display:flex; align-items:center; justify-content:center; gap:8px; padding:8px 16px; }
        .llevar-btn.occupied { border-color:var(--mustard); background:#3a2f1c; }
        .llevar-btn.selected { border-color:var(--chili); box-shadow:0 0 0 2px var(--chili); }
        .num-tables-ctrl { display:flex; align-items:center; gap:6px; margin-left:auto; background:var(--steel-800);
          border:1px solid var(--steel-700); border-radius:10px; padding:6px 10px; font-size:12px; color:var(--muted); }
        .num-tables-ctrl button { background:var(--steel-700); border:none; color:var(--cream); width:26px; height:26px;
          border-radius:6px; cursor:pointer; font-weight:700; }

        .work-grid { display:flex; gap:22px; flex-wrap:wrap; align-items:flex-start; margin-top:18px; }
        .catalog { flex:2; min-width:340px; }
        .cats { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:14px; }
        .cat-chip { background:var(--steel-800); border:1px solid var(--steel-700); color:var(--cream);
          padding:11px 18px; border-radius:22px; white-space:nowrap; cursor:pointer; font-size:14px; font-weight:700; min-height:44px; }
        .cat-chip.active { background:var(--mustard); color:var(--ink); border-color:var(--mustard); }

        .product-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(128px,1fr)); gap:10px; }
        .product-card { background:var(--steel-800); border:1px solid var(--steel-700); border-radius:11px; padding:11px 12px;
          cursor:pointer; transition:transform .12s, border-color .12s; min-height:64px; display:flex; flex-direction:column; justify-content:space-between; }
        .product-card:hover, .product-card:active { border-color:var(--mustard); transform:translateY(-2px); }
        .product-card .pname { font-weight:600; font-size:13px; margin-bottom:8px; line-height:1.25; }
        .product-card .pprice { font-family:'IBM Plex Mono',monospace; color:var(--mustard); font-size:13.5px; font-weight:700; }

        .ticket-wrap { flex:1; min-width:310px; position:sticky; top:20px; }
        .ticket { background:var(--paper); color:var(--ink); border-radius:8px 8px 0 0; padding:22px 22px 36px;
          clip-path:${zigzag}; box-shadow:0 8px 24px rgba(0,0,0,0.35); }
        .ticket h3 { font-family:'Bebas Neue',sans-serif; font-size:24px; margin:0 0 4px; letter-spacing:0.05em; }
        .ticket .sub { font-family:'IBM Plex Mono',monospace; font-size:12px; color:#7a715c; border-bottom:1px dashed var(--paper-line);
          padding-bottom:12px; margin-bottom:12px; }
        .ticket-line { display:flex; align-items:center; gap:9px; font-family:'IBM Plex Mono',monospace; font-size:14px;
          padding:8px 0; border-bottom:1px dotted var(--paper-line); }
        .ticket-line .lname { flex:1; }
        .qty-btn { width:26px; height:26px; border-radius:6px; border:1px solid var(--paper-line); background:#fff;
          display:flex; align-items:center; justify-content:center; cursor:pointer; }
        .ticket-total { display:flex; justify-content:space-between; align-items:baseline; margin-top:14px; }
        .ticket-total .lbl { font-family:'Bebas Neue',sans-serif; font-size:20px; letter-spacing:0.05em; }
        .ticket-total .amt { font-family:'IBM Plex Mono',monospace; font-size:24px; font-weight:700; }
        .charge-btn { width:100%; margin-top:16px; background:var(--chili); color:#fff; border:none; padding:16px;
          border-radius:10px; font-weight:700; font-size:16px; cursor:pointer; letter-spacing:0.03em; min-height:52px; }
        .charge-btn:disabled { background:#c8bfa8; cursor:not-allowed; }
        .charge-btn:hover:not(:disabled) { background:var(--chili-dark); }
        .pay-row { display:flex; gap:8px; margin-top:10px; }
        .pay-btn { flex:1; border:none; padding:14px 8px; border-radius:10px; font-weight:700; font-size:14px;
          cursor:pointer; min-height:50px; }
        .pay-btn:disabled { opacity:0.4; cursor:not-allowed; }
        .pay-cash { background:var(--green); color:#fff; }
        .pay-cash:hover:not(:disabled) { background:#3f7f50; }
        .pay-yape { background:#6739b7; color:#fff; }
        .pay-yape:hover:not(:disabled) { background:#542d94; }
        .clear-btn { width:100%; margin-top:8px; background:transparent; color:#8a8064; border:1px dashed #b8ac8a; padding:10px;
          border-radius:8px; font-weight:600; font-size:13px; cursor:pointer; display:flex; align-items:center; justify-content:center; gap:6px; }
        .empty-ticket { text-align:center; color:#8a8064; font-size:13px; padding:34px 0; }
        .flash-ok { position:fixed; bottom:24px; right:24px; background:var(--green); color:#fff; padding:14px 22px;
          border-radius:10px; font-weight:700; box-shadow:0 6px 20px rgba(0,0,0,0.3); z-index:50; }

        .filters { display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin-bottom:16px; }
        .quick-btn { background:var(--steel-800); border:1px solid var(--steel-700); color:var(--cream); padding:10px 16px;
          border-radius:9px; cursor:pointer; font-size:14px; font-weight:700; min-height:44px; }
        .date-input { background:var(--steel-800); border:1px solid var(--steel-700); color:var(--cream); padding:10px 12px;
          border-radius:9px; font-size:14px; min-height:44px; }
        .export-btn { background:var(--mustard); color:var(--ink); border:none; padding:12px 20px; border-radius:9px;
          font-weight:700; cursor:pointer; display:flex; align-items:center; gap:8px; margin-left:auto; min-height:44px; }
        .export-btn:hover { background:var(--mustard-dark); }

        .stat-cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:12px; margin-bottom:20px; }
        .stat-card { background:var(--steel-800); border:1px solid var(--steel-700); border-radius:14px; padding:18px; }
        .stat-card .k { font-size:12px; color:var(--muted); margin-bottom:6px; }
        .stat-card .v { font-family:'IBM Plex Mono',monospace; font-size:23px; font-weight:700; color:var(--mustard); }

        table { width:100%; border-collapse:collapse; background:var(--steel-800); border-radius:14px; overflow:hidden; }
        th, td { text-align:left; padding:12px 16px; font-size:14px; border-bottom:1px solid var(--steel-700); }
        th { color:var(--muted); font-weight:700; text-transform:uppercase; font-size:11px; letter-spacing:0.04em; }
        td.mono-cell { font-family:'IBM Plex Mono',monospace; }
        .del-icon { cursor:pointer; color:var(--muted); }
        .del-icon:hover { color:var(--chili); }

        .menu-toolbar { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:16px; }
        .menu-list { display:flex; flex-direction:column; gap:8px; margin-bottom:24px; }
        .menu-row { display:flex; align-items:center; gap:10px; background:var(--steel-800); border:1px solid var(--steel-700);
          border-radius:10px; padding:12px 16px; }
        .menu-row input[type=text] { flex:2; background:transparent; border:none; color:var(--cream); font-size:14px; font-weight:600; }
        .menu-row .cat-tag { font-size:11px; color:var(--muted); background:var(--steel-900); padding:5px 9px; border-radius:6px; min-width:130px; text-align:center; }
        .menu-row input[type=number] { width:90px; background:var(--steel-900); border:1px solid var(--steel-700); color:var(--mustard);
          font-family:'IBM Plex Mono',monospace; padding:7px 9px; border-radius:6px; }
        .add-form { display:flex; gap:8px; flex-wrap:wrap; background:var(--steel-800); border:1px dashed var(--steel-600);
          border-radius:14px; padding:18px; }
        .add-form input, .add-form select { background:var(--steel-900); border:1px solid var(--steel-700); color:var(--cream);
          padding:11px 14px; border-radius:9px; font-size:14px; min-height:44px; }
        .add-form input[name=name] { flex:2; min-width:180px; }
        .add-form button { background:var(--mustard); color:var(--ink); border:none; padding:11px 20px; border-radius:9px;
          font-weight:700; cursor:pointer; min-height:44px; }

        .notes-input { width:100%; background:#fff; border:1px solid var(--paper-line); color:var(--ink); border-radius:6px;
          padding:8px 10px; font-size:12px; font-family:'Inter',sans-serif; margin:10px 0 2px; resize:none; }
        .print-btn { width:100%; margin-top:8px; background:var(--steel-700); color:#fff; border:none; padding:12px;
          border-radius:8px; font-weight:700; font-size:13px; cursor:pointer; display:flex; align-items:center; justify-content:center; gap:7px; min-height:46px; }
        .print-btn:hover:not(:disabled) { background:var(--steel-600); }
        .print-btn:disabled { opacity:0.4; cursor:not-allowed; }
        .icon-btn { background:var(--steel-700); border:none; color:var(--cream); width:36px; height:36px; border-radius:8px;
          cursor:pointer; display:flex; align-items:center; justify-content:center; }
        .icon-btn:hover { background:var(--chili); }
        .open-note { font-size:12px; color:var(--mustard-dark); background:#3a2f1c; border-radius:6px; padding:6px 9px; margin-bottom:10px; }

        .open-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(280px,1fr)); gap:16px; }
        .open-card { background:var(--steel-800); border:1px solid var(--steel-700); border-radius:14px; padding:18px; display:flex; flex-direction:column; }
        .open-card-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:6px; }
        .open-card-head h4 { font-family:'Bebas Neue',sans-serif; font-size:20px; margin:0; letter-spacing:0.04em; color:var(--mustard); }
        .open-elapsed { font-size:12px; color:var(--muted); display:flex; align-items:center; gap:4px; margin-bottom:10px; }
        .open-items { flex:1; margin-bottom:12px; }
        .open-item-line { display:flex; justify-content:space-between; font-family:'IBM Plex Mono',monospace; font-size:13px;
          padding:4px 0; border-bottom:1px dotted var(--steel-700); }
        .open-total-row { display:flex; justify-content:space-between; align-items:baseline; margin:8px 0 12px; }
        .open-total-row .lbl { font-size:12px; color:var(--muted); text-transform:uppercase; letter-spacing:0.04em; }
        .open-total-row .amt { font-family:'IBM Plex Mono',monospace; font-size:20px; font-weight:700; color:var(--mustard); }
        .open-actions { display:flex; gap:8px; }
        .btn-secondary { flex:1; background:var(--steel-700); color:var(--cream); border:none; padding:11px; border-radius:8px;
          font-weight:600; font-size:13px; cursor:pointer; display:flex; align-items:center; justify-content:center; gap:6px; min-height:44px; }
        .btn-secondary:hover { background:var(--steel-600); }
        .btn-primary-open { flex:1.3; background:var(--chili); color:#fff; border:none; padding:11px; border-radius:8px;
          font-weight:700; font-size:13px; cursor:pointer; min-height:44px; }
        .btn-primary-open:hover { background:var(--chili-dark); }
        .empty-open { text-align:center; color:var(--muted); padding:60px 20px; font-size:15px; }

        .print-area { display:none; }
        @media print {
          @page { size: 80mm auto; margin: 3mm; }
          .screen-only { display:none !important; }
          .print-area { display:block !important; font-family:'IBM Plex Mono',monospace; color:#000; width:100%; }
          .print-title { font-family:'Bebas Neue',sans-serif; font-size:22px; text-align:center; letter-spacing:0.05em; }
          .print-sub { text-align:center; font-size:13px; font-weight:700; margin-bottom:2px; }
          .print-meta { text-align:center; font-size:11px; margin-bottom:8px; }
          .print-area hr { border:none; border-top:1px dashed #000; margin:8px 0; }
          .print-line { font-size:15px; font-weight:700; padding:4px 0; }
          .print-qty { display:inline-block; min-width:34px; }
          .print-notes { font-size:13px; font-weight:700; }
          .print-footer { text-align:center; font-size:10px; margin-top:10px; }
        }
      `}</style>

      <div className="print-area">
        {printOrder && (
          <div>
            <div className="print-title">COMPACHO'S</div>
            <div className="print-sub">COMANDA · COCINA</div>
            <div className="print-meta">{printOrder.label} · {printOrder.date} {printOrder.time}</div>
            <hr />
            {printOrder.items.map((l, idx) => (
              <div className="print-line" key={idx}><span className="print-qty">{l.qty}x</span>{l.name}</div>
            ))}
            {printOrder.notes && (
              <>
                <hr />
                <div className="print-notes">Notas: {printOrder.notes}</div>
              </>
            )}
            <hr />
            <div className="print-footer">Impreso {new Date().toLocaleString('es-PE')}</div>
          </div>
        )}
      </div>

      <div className="screen-only">
      <div className="topbar">
        <div className="brand">
          <ChefHat size={30} color="var(--mustard)" />
          <div>
            <h1 className="display">COMPACHO'S</h1>
            <div className="tag">Sabores buenachos</div>
          </div>
        </div>
        <div className="tabs">
          <button className={`tab-btn ${tab === 'pedido' ? 'active' : ''}`} onClick={() => setTab('pedido')}><UtensilsCrossed size={17}/> Tomar Pedido</button>
          <button className={`tab-btn ${tab === 'abiertos' ? 'active' : ''}`} onClick={() => setTab('abiertos')}>
            <Receipt size={17}/> Pedidos Abiertos {openKeys.length > 0 && <span className="badge-count">{openKeys.length}</span>}
          </button>
          <button className={`tab-btn ${tab === 'historial' ? 'active' : ''}`} onClick={() => setTab('historial')}><ClipboardList size={17}/> Historial</button>
          <button className={`tab-btn ${tab === 'menu' ? 'active' : ''}`} onClick={() => setTab('menu')}><Table2 size={17}/> Menú</button>
        </div>
        <span style={{ fontSize: 11, color: 'var(--green)' }}>● En línea</span>
      </div>

      <div className="content">
        {tab === 'pedido' && (
          <>
            <div className="section-label">Selecciona la mesa o para llevar</div>
            <div className="table-picker">
              <div className="table-grid">
                <button
                  className={`llevar-btn ${selectedKey === 'llevar' ? 'selected' : ''} ${openOrdersData['llevar']?.items?.length ? 'occupied' : ''}`}
                  onClick={() => setSelectedKey('llevar')}
                >
                  <ShoppingBag size={18}/> Para Llevar
                </button>
                {Array.from({ length: numTables }, (_, i) => i + 1).map((n) => {
                  const key = `mesa-${n}`;
                  const items = openOrdersData[key]?.items || [];
                  const total = items.reduce((s, l) => s + l.price * l.qty, 0);
                  return (
                    <button key={n} className={`table-btn ${selectedKey === key ? 'selected' : ''} ${items.length ? 'occupied' : ''}`}
                      onClick={() => setSelectedKey(key)}>
                      <span>Mesa {n}</span>
                      {items.length > 0 && <span className="amt">{fmt(total)}</span>}
                    </button>
                  );
                })}
              </div>
              <div className="num-tables-ctrl">
                N° mesas
                <button onClick={() => changeNumTables(numTables - 1)}>–</button>
                <span>{numTables}</span>
                <button onClick={() => changeNumTables(numTables + 1)}>+</button>
              </div>
            </div>

            <div className="work-grid">
              <div className="catalog">
                <div className="cats">
                  {CATEGORIES.map((c) => (
                    <button key={c} className={`cat-chip ${activeCat === c ? 'active' : ''}`} onClick={() => setActiveCat(c)}>{c}</button>
                  ))}
                </div>
                <div className="product-grid">
                  {menu.filter((m) => m.category === activeCat).map((item) => (
                    <div key={item.id} className="product-card" onClick={() => addToCart(item)}>
                      <div className="pname">{item.name}</div>
                      <div className="pprice">{fmt(item.price)}</div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="ticket-wrap">
                <div className="ticket">
                  <h3>Comanda</h3>
                  <div className="sub">{isTable ? `MESA ${tableNum}` : 'PARA LLEVAR'} · {new Date().toLocaleDateString('es-PE')}</div>
                  {cart.length === 0 && <div className="empty-ticket">Toca un producto del menú<br/>para agregarlo al pedido</div>}
                  {cart.map((l) => (
                    <div className="ticket-line" key={l.menuId}>
                      <button className="qty-btn" onClick={() => changeQty(l.menuId, -1)}><Minus size={13}/></button>
                      <span>{l.qty}</span>
                      <button className="qty-btn" onClick={() => changeQty(l.menuId, 1)}><Plus size={13}/></button>
                      <span className="lname">{l.name}</span>
                      <span>{fmt(l.price * l.qty)}</span>
                      <X size={16} style={{cursor:'pointer'}} onClick={() => removeLine(l.menuId)} />
                    </div>
                  ))}
                  {cart.length > 0 && (
                    <div className="ticket-total">
                      <span className="lbl">TOTAL</span>
                      <span className="amt">{fmt(cartTotal)}</span>
                    </div>
                  )}
                  {cart.length > 0 && (
                    <textarea className="notes-input" rows={2} placeholder="Notas para cocina (opcional): sin cebolla, término medio, etc."
                      value={cartData.notes || ''}
                      onChange={(e) => updateNotes(e.target.value)}
                      onBlur={commitNotes} />
                  )}
                  <button className="print-btn" disabled={cart.length === 0} onClick={() => printComanda(selectedKey)}>
                    <Printer size={15}/> Imprimir comanda
                  </button>
                  <div className="pay-row">
                    <button className="pay-btn pay-cash" disabled={cart.length === 0} onClick={() => chargeOrder('efectivo')}>
                      💵 Efectivo
                    </button>
                    <button className="pay-btn pay-yape" disabled={cart.length === 0} onClick={() => chargeOrder('yape')}>
                      📱 Yape
                    </button>
                  </div>
                  {cart.length > 0 && (
                    <button className="clear-btn" onClick={clearCart}><RotateCcw size={13}/> Vaciar pedido sin cobrar</button>
                  )}
                </div>
              </div>
            </div>
            {confirmFlash && <div className="flash-ok">Pedido cobrado ✓</div>}
          </>
        )}

        {tab === 'abiertos' && (
          <>
            <div className="section-label">Mesas y pedidos activos ({openKeys.length})</div>
            {openKeys.length === 0 && (
              <div className="empty-open">No hay mesas ni pedidos abiertos en este momento.<br/>Ve a "Tomar Pedido" para iniciar uno.</div>
            )}
            <div className="open-grid">
              {openKeys.map((key) => {
                const data = openOrdersData[key];
                const items = data.items || [];
                const total = items.reduce((s, l) => s + l.price * l.qty, 0);
                const label = key === 'llevar' ? 'Para Llevar' : `Mesa ${key.replace('mesa-', '')}`;
                return (
                  <div className="open-card" key={key}>
                    <div className="open-card-head">
                      <h4>{label}</h4>
                      <button className="icon-btn" title="Imprimir comanda" onClick={() => printComanda(key)}><Printer size={16}/></button>
                    </div>
                    <div className="open-elapsed"><Clock size={13}/> Abierta hace {minutesAgo(data.opened_at)} min</div>
                    {data.notes && <div className="open-note">📝 {data.notes}</div>}
                    <div className="open-items">
                      {items.map((l) => (
                        <div className="open-item-line" key={l.menuId}>
                          <span>{l.qty}x {l.name}</span>
                          <span>{fmt(l.price * l.qty)}</span>
                        </div>
                      ))}
                    </div>
                    <div className="open-total-row">
                      <span className="lbl">Total</span>
                      <span className="amt">{fmt(total)}</span>
                    </div>
                    <div className="open-actions">
                      <button className="btn-secondary" onClick={() => { setSelectedKey(key); setTab('pedido'); }}>
                        <ArrowRight size={14}/> Agregar más
                      </button>
                    </div>
                    <div className="pay-row">
                      <button className="pay-btn pay-cash" onClick={() => chargeOrderForKey(key, 'efectivo')}>💵 Efectivo</button>
                      <button className="pay-btn pay-yape" onClick={() => chargeOrderForKey(key, 'yape')}>📱 Yape</button>
                    </div>
                  </div>
                );
              })}
            </div>
            {confirmFlash && <div className="flash-ok">Pedido cobrado ✓</div>}
          </>
        )}

        {tab === 'historial' && (
          <>
            <div className="filters">
              <button className="quick-btn" onClick={() => { setDateFrom(todayISO()); setDateTo(todayISO()); }}>Hoy</button>
              <button className="quick-btn" onClick={() => { setDateFrom(weekAgoISO()); setDateTo(todayISO()); }}>Últimos 7 días</button>
              <button className="quick-btn" onClick={() => { setDateFrom('2000-01-01'); setDateTo(todayISO()); }}>Todo</button>
              <input className="date-input" type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} />
              <span style={{color:'var(--muted)'}}>a</span>
              <input className="date-input" type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
              <button className="export-btn" onClick={exportExcel}><Download size={16}/> Exportar a Excel</button>
            </div>

            <div className="stat-cards">
              <div className="stat-card"><div className="k">Ventas totales</div><div className="v">{fmt(totalVentas)}</div></div>
              <div className="stat-card"><div className="k">N° de pedidos</div><div className="v">{filteredOrders.length}</div></div>
              <div className="stat-card"><div className="k">Ticket promedio</div><div className="v">{fmt(ticketProm)}</div></div>
              <div className="stat-card"><div className="k">Para llevar / Mesa</div><div className="v" style={{fontSize:16}}>{fmt(ventasLlevar)} / {fmt(ventasMesa)}</div></div>
              <div className="stat-card"><div className="k">Efectivo / Yape</div><div className="v" style={{fontSize:16}}>{fmt(ventasEfectivo)} / {fmt(ventasYape)}</div></div>
            </div>

            <table>
              <thead><tr><th>Hora</th><th>Tipo</th><th>Pago</th><th>Productos</th><th>Total</th><th></th></tr></thead>
              <tbody>
                {filteredOrders.map((o) => (
                  <tr key={o.id}>
                    <td className="mono-cell">{o.date} {o.time}</td>
                    <td>{o.type === 'llevar' ? 'Para llevar' : `Mesa ${o.table}`}</td>
                    <td>{o.payment === 'yape' ? '📱 Yape' : '💵 Efectivo'}</td>
                    <td>{o.items.map((i) => `${i.qty}x ${i.name}`).join(', ')}</td>
                    <td className="mono-cell">{fmt(o.total)}</td>
                    <td><Trash2 size={16} className="del-icon" onClick={() => deleteSale(o.id)} /></td>
                  </tr>
                ))}
                {filteredOrders.length === 0 && (
                  <tr><td colSpan="6" style={{textAlign:'center', color:'var(--muted)', padding:24}}>No hay pedidos en este rango de fechas.</td></tr>
                )}
              </tbody>
            </table>
          </>
        )}

        {tab === 'menu' && (
          <>
            <div className="menu-toolbar">
              <button className={`cat-chip ${menuCat === 'Todos' ? 'active' : ''}`} onClick={() => setMenuCat('Todos')}>Todos</button>
              {CATEGORIES.map((c) => (
                <button key={c} className={`cat-chip ${menuCat === c ? 'active' : ''}`} onClick={() => setMenuCat(c)}>{c}</button>
              ))}
            </div>
            <div className="menu-list">
              {menuByCat.map((it) => (
                <div className="menu-row" key={it.id}>
                  <input type="text" value={it.name} onChange={(e) => updateMenuItem(it.id, 'name', e.target.value)} />
                  <span className="cat-tag">{it.category}</span>
                  <span style={{color:'var(--muted)', fontSize:13}}>S/</span>
                  <input type="number" step="0.5" value={it.price} onChange={(e) => updateMenuItem(it.id, 'price', e.target.value)} />
                  <Trash2 size={17} className="del-icon" onClick={() => deleteMenuItem(it.id)} />
                </div>
              ))}
            </div>
            <div className="add-form">
              <input name="name" type="text" placeholder="Nombre del producto" value={newProd.name}
                onChange={(e) => setNewProd({ ...newProd, name: e.target.value })} />
              <select value={newProd.category} onChange={(e) => setNewProd({ ...newProd, category: e.target.value })}>
                {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
              <input type="number" step="0.5" placeholder="Precio" value={newProd.price}
                onChange={(e) => setNewProd({ ...newProd, price: e.target.value })} />
              <button onClick={addProduct}>+ Agregar producto</button>
            </div>
          </>
        )}
      </div>
      </div>
    </div>
  );
}

