#!/usr/bin/env python3
"""
DZAIR Sales & Profit Manager - Improved Version
Features added:
- Improved UI layout and basic form validation
- Embedded Matplotlib charts in the Dashboard
- Advanced Filters: search and date range filtering
- Export & backup (Excel, PDF)
Notes:
- Requires: pandas, matplotlib, reportlab
- To build .exe: see build_exe.bat and README.md instructions included in package.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3, os, math
from datetime import datetime, timedelta
import pandas as pd

# Matplotlib imports for embedding in Tkinter
try:
    import matplotlib
    matplotlib.use("Agg")  # safe backend for headless, embedding will switch when needed
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False

DB_PATH = os.path.join(os.path.dirname(__file__), "dzair.db")

# ---- Calculations ----
def tot_livraison(weight, delivery_price):
    return (weight * 50) + delivery_price

def p_fayda(selling_price, tot_livraison, purchase_price):
    return (selling_price - tot_livraison) - purchase_price

def fayda_safia(pf):
    return pf - 500

# ---- DB ----
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS Clients (
        client_id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT, address TEXT, city TEXT,
        total_spent REAL DEFAULT 0, total_orders INTEGER DEFAULT 0)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS Products (
        product_id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, purchase_price REAL NOT NULL, weight REAL,
        default_delivery_price REAL DEFAULT 0, selling_price REAL, stock_qty INTEGER DEFAULT 0)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS Sales (
        sale_id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT UNIQUE, client_id INTEGER, product_id INTEGER,
        quantity INTEGER DEFAULT 1, purchase_price REAL, selling_price REAL, weight REAL, delivery_price REAL,
        tot_livraison REAL, p_fayda REAL, fayda_safia REAL, payment_method TEXT, status TEXT, paid INTEGER DEFAULT 0, date TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS SponsoredFees (
        fee_id INTEGER PRIMARY KEY AUTOINCREMENT, campaign_name TEXT NOT NULL, platform TEXT, amount_spent REAL NOT NULL, date TEXT)""")
    conn.commit()
    conn.close()

def generate_invoice_no(conn, date_str=None):
    if date_str is None:
        date_str = datetime.now().strftime("%Y-%m-%d")
    year = datetime.fromisoformat(date_str).year
    cur = conn.cursor()
    cur.execute("SELECT invoice_no FROM Sales WHERE invoice_no LIKE ? ORDER BY sale_id DESC LIMIT 1", (f'DZAIR-{year}-%',))
    row = cur.fetchone()
    if row is None:
        seq = 1
    else:
        last = row[0]
        try:
            seq = int(last.split('-')[-1]) + 1
        except:
            seq = 1
    return f"DZAIR-{year}-{seq:03d}"

# ---- App ----
class App(ttk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.root = root
        root.title("DZAIR - Sales & Profit Manager (Improved)")
        root.geometry("1100x700")
        self.style = ttk.Style(root)
        # Try to use a modern theme if available
        for theme in ("clam","alt","default"):
            try:
                self.style.theme_use(theme); break
            except Exception:
                pass
        # larger fonts for readability
        default_font = ("Segoe UI", 10)
        root.option_add("*Font", default_font)

        init_db()
        self.conn = get_conn()
        self.create_widgets()
        self.refresh_all()

    def create_widgets(self):
        # Top toolbar for quick actions & filters
        toolbar = ttk.Frame(self.root, padding=6)
        toolbar.pack(side="top", fill="x")

        ttk.Button(toolbar, text="Add Client", command=lambda: self.open_tab("Clients")).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Add Product", command=lambda: self.open_tab("Products")).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Add Sale", command=lambda: self.open_tab("Sales")).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Dashboard", command=lambda: self.open_tab("Dashboard")).pack(side="left", padx=4)
        ttk.Button(toolbar, text="Sponsored Fees", command=lambda: self.open_tab("Fees")).pack(side="left", padx=4)

        # Search and date filters
        ttk.Label(toolbar, text="Search:").pack(side="left", padx=(20,4))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=30)
        search_entry.pack(side="left")
        search_entry.bind("<Return>", lambda e: self.refresh_sales())

        ttk.Label(toolbar, text="From:").pack(side="left", padx=(10,4))
        self.date_from = ttk.Entry(toolbar, width=12)
        self.date_from.pack(side="left")
        ttk.Label(toolbar, text="To:").pack(side="left", padx=(6,4))
        self.date_to = ttk.Entry(toolbar, width=12)
        self.date_to.pack(side="left")
        ttk.Button(toolbar, text="Apply Filters", command=self.refresh_sales).pack(side="left", padx=6)

        # Notebook
        self.nb = ttk.Notebook(self.root); self.nb.pack(fill="both", expand=True)
        # Tabs
        self.tab_clients = ttk.Frame(self.nb); self.nb.add(self.tab_clients, text="Clients")
        self.tab_products = ttk.Frame(self.nb); self.nb.add(self.tab_products, text="Products")
        self.tab_sales = ttk.Frame(self.nb); self.nb.add(self.tab_sales, text="Sales")
        self.tab_dashboard = ttk.Frame(self.nb); self.nb.add(self.tab_dashboard, text="Dashboard")
        self.tab_fees = ttk.Frame(self.nb); self.nb.add(self.tab_fees, text="Sponsored Fees")
        self.tab_reports = ttk.Frame(self.nb); self.nb.add(self.tab_reports, text="Reports / Export")

        self.build_clients_tab(); self.build_products_tab(); self.build_sales_tab(); self.build_dashboard_tab(); self.build_fees_tab(); self.build_reports_tab()

    def open_tab(self, name):
        mapping = {"Clients":0, "Products":1, "Sales":2, "Dashboard":3, "Fees":4, "Reports":5}
        idx = mapping.get(name, 0); self.nb.select(idx)

    # ------------- Clients -------------
    def build_clients_tab(self):
        frame = self.tab_clients
        form = ttk.Frame(frame, padding=8)
        form.pack(side="left", fill="y", padx=10, pady=10)
        ttk.Label(form, text="Name *").pack(anchor="w"); self.c_name = ttk.Entry(form); self.c_name.pack(fill="x")
        ttk.Label(form, text="Phone").pack(anchor="w"); self.c_phone = ttk.Entry(form); self.c_phone.pack(fill="x")
        ttk.Label(form, text="Address").pack(anchor="w"); self.c_address = ttk.Entry(form); self.c_address.pack(fill="x")
        ttk.Label(form, text="City").pack(anchor="w"); self.c_city = ttk.Entry(form); self.c_city.pack(fill="x")
        ttk.Button(form, text="Add Client", command=self.add_client).pack(pady=6)
        ttk.Button(form, text="Clear", command=lambda: [self.c_name.delete(0,'end'), self.c_phone.delete(0,'end'), self.c_address.delete(0,'end'), self.c_city.delete(0,'end')]).pack()

        list_frame = ttk.Frame(frame); list_frame.pack(fill="both", expand=True, padx=8, pady=8)
        cols = ("client_id","name","phone","address","city","total_spent","total_orders")
        self.clients_tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="browse")
        for c in cols: self.clients_tree.heading(c, text=c)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.clients_tree.yview); self.clients_tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y"); self.clients_tree.pack(fill="both", expand=True)

    def add_client(self):
        name = self.c_name.get().strip()
        if not name:
            messagebox.showerror("Validation", "Client name is required."); return
        cur = self.conn.cursor()
        cur.execute("INSERT INTO Clients (name, phone, address, city) VALUES (?, ?, ?, ?)", (name, self.c_phone.get().strip(), self.c_address.get().strip(), self.c_city.get().strip()))
        self.conn.commit(); messagebox.showinfo("OK","Client added.")
        self.c_name.delete(0,'end'); self.c_phone.delete(0,'end'); self.c_address.delete(0,'end'); self.c_city.delete(0,'end')
        self.refresh_clients()

    def refresh_clients(self):
        for r in self.clients_tree.get_children(): self.clients_tree.delete(r)
        cur = self.conn.cursor()
        for row in cur.execute("SELECT * FROM Clients ORDER BY client_id DESC"): self.clients_tree.insert("", "end", values=tuple(row))

    # ------------- Products -------------
    def build_products_tab(self):
        frame = self.tab_products
        form = ttk.Frame(frame, padding=8); form.pack(side="left", fill="y", padx=10, pady=10)
        ttk.Label(form, text="Product Name *").pack(anchor="w"); self.p_name = ttk.Entry(form); self.p_name.pack(fill="x")
        ttk.Label(form, text="Purchase Price *").pack(anchor="w"); self.p_purchase = ttk.Entry(form); self.p_purchase.pack(fill="x")
        ttk.Label(form, text="Weight (kg)").pack(anchor="w"); self.p_weight = ttk.Entry(form); self.p_weight.pack(fill="x")
        ttk.Label(form, text="Default Delivery").pack(anchor="w"); self.p_del = ttk.Entry(form); self.p_del.pack(fill="x")
        ttk.Label(form, text="Selling Price *").pack(anchor="w"); self.p_sell = ttk.Entry(form); self.p_sell.pack(fill="x")
        ttk.Label(form, text="Stock Qty").pack(anchor="w"); self.p_stock = ttk.Entry(form); self.p_stock.pack(fill="x")
        ttk.Button(form, text="Add Product", command=self.add_product).pack(pady=6)
        ttk.Button(form, text="Clear", command=lambda: [self.p_name.delete(0,'end'), self.p_purchase.delete(0,'end'), self.p_weight.delete(0,'end'), self.p_del.delete(0,'end'), self.p_sell.delete(0,'end'), self.p_stock.delete(0,'end')]).pack()

        list_frame = ttk.Frame(frame); list_frame.pack(fill="both", expand=True, padx=8, pady=8)
        cols = ("product_id","name","purchase_price","weight","default_delivery_price","selling_price","stock_qty")
        self.products_tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="browse")
        for c in cols: self.products_tree.heading(c, text=c)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.products_tree.yview); self.products_tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y"); self.products_tree.pack(fill="both", expand=True)

    def add_product(self):
        try:
            purchase = float(self.p_purchase.get()); selling = float(self.p_sell.get())
            weight = float(self.p_weight.get() or 0); default_del = float(self.p_del.get() or 0); stock = int(self.p_stock.get() or 0)
        except Exception:
            messagebox.showerror("Validation", "Please enter numeric values for prices, weight and stock."); return
        name = self.p_name.get().strip()
        if not name:
            messagebox.showerror("Validation", "Product name required."); return
        cur = self.conn.cursor()
        cur.execute("INSERT INTO Products (name, purchase_price, weight, default_delivery_price, selling_price, stock_qty) VALUES (?, ?, ?, ?, ?, ?)",
                    (name, purchase, weight, default_del, selling, stock))
        self.conn.commit(); messagebox.showinfo("OK","Product added.")
        for e in (self.p_name, self.p_purchase, self.p_weight, self.p_del, self.p_sell, self.p_stock): e.delete(0,'end')
        self.refresh_products()

    def refresh_products(self):
        for r in self.products_tree.get_children(): self.products_tree.delete(r)
        cur = self.conn.cursor()
        for row in cur.execute("SELECT * FROM Products ORDER BY product_id DESC"): self.products_tree.insert("", "end", values=tuple(row))

    # ------------- Sales -------------
    def build_sales_tab(self):
        frame = self.tab_sales
        top = ttk.Frame(frame, padding=8); top.pack(fill="x")
        # client/product selectors
        ttk.Label(top, text="Client").grid(row=0,column=0, sticky="w"); self.sale_client = ttk.Combobox(top, width=40); self.sale_client.grid(row=0,column=1, sticky="w")
        ttk.Label(top, text="Product").grid(row=1,column=0, sticky="w"); self.sale_product = ttk.Combobox(top, width=40); self.sale_product.grid(row=1,column=1, sticky="w")
        ttk.Label(top, text="Qty").grid(row=0,column=2, sticky="w"); self.sale_qty = ttk.Entry(top, width=8); self.sale_qty.grid(row=0,column=3, sticky="w")
        ttk.Label(top, text="Delivery (opt)").grid(row=1,column=2, sticky="w"); self.sale_delivery = ttk.Entry(top, width=8); self.sale_delivery.grid(row=1,column=3, sticky="w")
        ttk.Label(top, text="Payment").grid(row=2,column=0, sticky="w"); self.sale_payment = ttk.Combobox(top, values=["Cash","BaridiMob","CCP","Bank"]); self.sale_payment.grid(row=2,column=1, sticky="w")
        ttk.Label(top, text="Status").grid(row=2,column=2, sticky="w"); self.sale_status = ttk.Combobox(top, values=["Pending","Delivered","Returned"]); self.sale_status.grid(row=2,column=3, sticky="w")
        ttk.Button(top, text="Add Sale", command=self.add_sale).grid(row=3,column=0, columnspan=2, pady=6)
        ttk.Button(top, text="Clear", command=lambda: [self.sale_client.set(''), self.sale_product.set(''), self.sale_qty.delete(0,'end'), self.sale_delivery.delete(0,'end')]).grid(row=3, column=2, columnspan=2)

        # Sales table and controls
        mid = ttk.Frame(frame); mid.pack(fill="both", expand=True, padx=8, pady=8)
        cols = ("sale_id","invoice_no","date","client","product","qty","selling_price","tot_livraison","p_fayda","fayda_safia","status")
        self.sales_tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        for c in cols: self.sales_tree.heading(c, text=c)
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.sales_tree.yview"); self.sales_tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y"); self.sales_tree.pack(fill="both", expand=True)

        # bottom buttons
        bottom = ttk.Frame(frame); bottom.pack(fill="x", padx=8, pady=6)
        ttk.Button(bottom, text="Refresh", command=self.refresh_sales).pack(side="left")
        ttk.Button(bottom, text="Export Filtered to Excel", command=self.export_filtered_excel).pack(side="left", padx=6)
        ttk.Button(bottom, text="Export Selected Invoice to PDF", command=self.export_invoice_pdf).pack(side="left", padx=6)

    def refresh_sales(self):
        # load client/product maps
        cur = self.conn.cursor()
        cur.execute("SELECT client_id, name FROM Clients ORDER BY name"); clients = cur.fetchall()
        self.client_map = {f\"{r['name']} ({r['client_id']})\": r['client_id'] for r in clients}
        self.sale_client['values'] = list(self.client_map.keys())
        cur.execute("SELECT product_id, name FROM Products ORDER BY name"); products = cur.fetchall()
        self.product_map = {f\"{r['name']} ({r['product_id']})\": r['product_id'] for r in products}
        self.sale_product['values'] = list(self.product_map.keys())

        # Build query with filters
        q = \"\"\"SELECT S.*, C.name as client_name, P.name as product_name FROM Sales S
                 LEFT JOIN Clients C ON S.client_id=C.client_id
                 LEFT JOIN Products P ON S.product_id=P.product_id WHERE 1=1\"\"\"
        params = []
        s = self.search_var.get().strip()
        if s:
            q += \" AND (C.name LIKE ? OR P.name LIKE ? OR S.invoice_no LIKE ? OR S.status LIKE ?)\")
            sparam = f\"%{s}%\"
            params.extend([sparam, sparam, sparam, sparam])
        # date range filters in YYYY-MM-DD - if empty ignore
        df = self.date_from.get().strip()
        dt = self.date_to.get().strip()
        if df:
            try:
                datetime.fromisoformat(df)
                q += \" AND date(S.date) >= date(?)\"; params.append(df)
            except Exception:
                messagebox.showerror(\"Date Error\", \"From date must be YYYY-MM-DD\"); return
        if dt:
            try:
                datetime.fromisoformat(dt)
                q += \" AND date(S.date) <= date(?)\"; params.append(dt)
            except Exception:
                messagebox.showerror(\"Date Error\", \"To date must be YYYY-MM-DD\"); return
        q += \" ORDER BY S.sale_id DESC\"
        # execute and populate tree
        for r in self.sales_tree.get_children(): self.sales_tree.delete(r)
        cur = self.conn.cursor()
        for row in cur.execute(q, params):
            self.sales_tree.insert('', 'end', values=(row['sale_id'], row['invoice_no'], row['date'], row['client_name'], row['product_name'], row['quantity'], row['selling_price'], row['tot_livraison'], row['p_fayda'], row['fayda_safia'], row['status']))

    def add_sale(self):
        client_key = self.sale_client.get(); prod_key = self.sale_product.get()
        if client_key not in self.client_map or prod_key not in self.product_map:
            messagebox.showerror("Validation", "Select valid client and product"); return
        try:
            qty = int(self.sale_qty.get() or 1)
        except Exception:
            messagebox.showerror("Validation", "Quantity must be integer"); return
        try:
            delivery_price = float(self.sale_delivery.get()) if self.sale_delivery.get().strip() else None
        except Exception:
            messagebox.showerror("Validation", "Delivery must be numeric"); return
        payment = self.sale_payment.get() or "Cash"
        status = self.sale_status.get() or "Pending"
        date_str = datetime.now().strftime("%Y-%m-%d")
        client_id = self.client_map[client_key]; product_id = self.product_map[prod_key]
        cur = self.conn.cursor()
        cur.execute("SELECT purchase_price, weight, default_delivery_price, selling_price, stock_qty FROM Products WHERE product_id=?", (product_id,))
        prod = cur.fetchone()
        if not prod: messagebox.showerror("Error", "Product not found"); return
        purchase_price = prod['purchase_price']; weight = prod['weight'] or 0; default_del = prod['default_delivery_price'] or 0; selling_price = prod['selling_price'] or 0
        if delivery_price is None: delivery_price = default_del
        tot_liv = tot_livraison(weight, delivery_price); pf = p_fayda(selling_price, tot_liv, purchase_price); fs = fayda_safia(pf)
        invoice = generate_invoice_no(self.conn, date_str)
        cur.execute("INSERT INTO Sales (invoice_no, client_id, product_id, quantity, purchase_price, selling_price, weight, delivery_price, tot_livraison, p_fayda, fayda_safia, payment_method, status, paid, date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (invoice, client_id, product_id, qty, purchase_price, selling_price, weight, delivery_price, tot_liv, pf, fs, payment, status, 1 if status=='Delivered' else 0, date_str))
        cur.execute("UPDATE Clients SET total_spent = total_spent + ?, total_orders = total_orders + 1 WHERE client_id = ?", (selling_price*qty, client_id))
        cur.execute("UPDATE Products SET stock_qty = stock_qty - ? WHERE product_id = ?", (qty, product_id))
        self.conn.commit(); messagebox.showinfo("Sale Added", f"Invoice: {invoice}"); self.refresh_sales()

    # ------------- Dashboard -------------
    def build_dashboard_tab(self):
        frame = self.tab_dashboard
        top = ttk.Frame(frame, padding=8); top.pack(fill="x")
        self.lbl_sales = ttk.Label(top, text="Sales: 0"); self.lbl_sales.pack(side="left", padx=6)
        self.lbl_profit = ttk.Label(top, text="Profit: 0"); self.lbl_profit.pack(side="left", padx=6)
        self.lbl_delivery = ttk.Label(top, text="Delivery: 0"); self.lbl_delivery.pack(side="left", padx=6)
        ttk.Button(top, text="Refresh", command=self.refresh_dashboard).pack(side="left", padx=6)
        # chart area
        self.chart_frame = ttk.Frame(frame); self.chart_frame.pack(fill="both", expand=True, padx=8, pady=8)
        if MATPLOTLIB_AVAILABLE:
            # create a matplotlib Figure
            self.fig = Figure(figsize=(6,4), dpi=100)
            self.ax = self.fig.add_subplot(111)
            self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
            self.canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            ttk.Label(self.chart_frame, text="Matplotlib not available - install matplotlib to see charts").pack()

    def refresh_dashboard(self):
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) FROM Sales"); sales_count = cur.fetchone()[0]
        cur.execute("SELECT SUM(p_fayda) FROM Sales"); total_profit = cur.fetchone()[0] or 0
        cur.execute("SELECT SUM(tot_livraison) FROM Sales"); total_delivery = cur.fetchone()[0] or 0
        self.lbl_sales.config(text=f"Sales: {sales_count}"); self.lbl_profit.config(text=f"Profit: {total_profit:.2f}"); self.lbl_delivery.config(text=f"Delivery: {total_delivery:.2f}")
        if MATPLOTLIB_AVAILABLE:
            # Sales per day (last 14 days)
            cur.execute("SELECT date, SUM(p_fayda) as profit, COUNT(*) as cnt FROM Sales WHERE date >= date('now','-13 days') GROUP BY date ORDER BY date")
            rows = cur.fetchall(); dates = [r[0] for r in rows]; profits = [r[1] for r in rows]
            self.ax.clear()
            if dates:
                self.ax.plot(dates, profits, marker='o')
                self.ax.set_title("Profit last 14 days"); self.ax.set_ylabel("Profit"); self.ax.set_xticks(dates[::max(1,len(dates)//7)])
                for label in self.ax.get_xticklabels(): label.set_rotation(30)
            else:
                self.ax.text(0.5, 0.5, "No data", ha='center', va='center')
            self.canvas.draw()

    # ------------- Sponsored Fees -------------
    def build_fees_tab(self):
        frame = self.tab_fees
        left = ttk.Frame(frame, padding=8); left.pack(side="left", fill="y", padx=8, pady=8)
        ttk.Label(left, text="Campaign").pack(anchor="w"); self.f_name = ttk.Entry(left); self.f_name.pack(fill="x")
        ttk.Label(left, text="Platform").pack(anchor="w"); self.f_platform = ttk.Entry(left); self.f_platform.pack(fill="x")
        ttk.Label(left, text="Amount").pack(anchor="w"); self.f_amount = ttk.Entry(left); self.f_amount.pack(fill="x")
        ttk.Button(left, text="Add Fee", command=self.add_fee).pack(pady=6)
        ttk.Button(left, text="Clear", command=lambda: [self.f_name.delete(0,'end'), self.f_platform.delete(0,'end'), self.f_amount.delete(0,'end')]).pack()

        list_frame = ttk.Frame(frame); list_frame.pack(fill="both", expand=True, padx=8, pady=8)
        cols = ("fee_id","campaign_name","platform","amount_spent","date")
        self.fees_tree = ttk.Treeview(list_frame, columns=cols, show="headings"); 
        for c in cols: self.fees_tree.heading(c, text=c)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.fees_tree.yview); self.fees_tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y"); self.fees_tree.pack(fill="both", expand=True)

    def add_fee(self):
        name = self.f_name.get().strip()
        if not name: messagebox.showerror("Validation", "Campaign name required"); return
        try: amt = float(self.f_amount.get()); 
        except: messagebox.showerror("Validation", "Amount must be numeric"); return
        plat = self.f_platform.get().strip(); date_str = datetime.now().strftime("%Y-%m-%d")
        cur = self.conn.cursor(); cur.execute("INSERT INTO SponsoredFees (campaign_name, platform, amount_spent, date) VALUES (?, ?, ?, ?)", (name, plat, amt, date_str)); self.conn.commit()
        messagebox.showinfo("OK","Fee added"); self.refresh_fees(); self.f_name.delete(0,'end'); self.f_platform.delete(0,'end'); self.f_amount.delete(0,'end')

    def refresh_fees(self):
        for r in self.fees_tree.get_children(): self.fees_tree.delete(r)
        cur = self.conn.cursor()
        for row in cur.execute("SELECT * FROM SponsoredFees ORDER BY fee_id DESC"): self.fees_tree.insert("", "end", values=tuple(row))

    # ------------- Reports & Export -------------
    def build_reports_tab(self):
        frame = self.tab_reports
        ttk.Button(frame, text="Export All Tables to Excel", command=self.export_all_excel).pack(pady=8, padx=8)
        ttk.Button(frame, text="Export Filtered Sales to Excel", command=self.export_filtered_excel).pack(pady=8, padx=8)
        ttk.Button(frame, text="Backup DB (Copy .db)", command=self.backup_db).pack(pady=8, padx=8)
        ttk.Button(frame, text="Open App Folder", command=lambda: os.startfile(os.path.dirname(__file__))).pack(pady=8, padx=8)

    def export_all_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not path: return
        df_clients = pd.read_sql_query("SELECT * FROM Clients", self.conn)
        df_products = pd.read_sql_query("SELECT * FROM Products", self.conn)
        df_sales = pd.read_sql_query("SELECT * FROM Sales", self.conn)
        df_fees = pd.read_sql_query("SELECT * FROM SponsoredFees", self.conn)
        with pd.ExcelWriter(path) as writer:
            df_clients.to_excel(writer, sheet_name="Clients", index=False)
            df_products.to_excel(writer, sheet_name="Products", index=False)
            df_sales.to_excel(writer, sheet_name="Sales", index=False)
            df_fees.to_excel(writer, sheet_name="SponsoredFees", index=False)
        messagebox.showinfo("Exported", f"Saved to {path}")

    def export_filtered_excel(self):
        # export current sales view (with filters applied)
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not path: return
        # Reuse query builder from refresh_sales but without populating UI
        q = "SELECT S.*, C.name as client_name, P.name as product_name FROM Sales S LEFT JOIN Clients C ON S.client_id=C.client_id LEFT JOIN Products P ON S.product_id=P.product_id WHERE 1=1"
        params = []
        s = self.search_var.get().strip()
        if s: q += " AND (C.name LIKE ? OR P.name LIKE ? OR S.invoice_no LIKE ? OR S.status LIKE ?)"; sparam = f"%{s}%"; params.extend([sparam,sparam,sparam,sparam])
        df = self.date_from.get().strip(); dt = self.date_to.get().strip()
        if df:
            try: datetime.fromisoformat(df); q += " AND date(S.date) >= date(?)"; params.append(df)
            except: messagebox.showerror("Date Error", "From date must be YYYY-MM-DD"); return
        if dt:
            try: datetime.fromisoformat(dt); q += " AND date(S.date) <= date(?)"; params.append(dt)
            except: messagebox.showerror("Date Error", "To date must be YYYY-MM-DD"); return
        q += " ORDER BY S.sale_id DESC"
        df_sales = pd.read_sql_query(q, self.conn, params=params)
        df_sales.to_excel(path, index=False)
        messagebox.showinfo("Exported", f"Filtered sales saved to {path}")

    def export_invoice_pdf(self):
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas
        except Exception:
            messagebox.showerror("Error", "reportlab not installed. Install with: pip install reportlab"); return
        sel = self.sales_tree.selection()
        if not sel: messagebox.showerror("Error", "Select a sale first"); return
        sale_id = self.sales_tree.item(sel[0])['values'][0]
        cur = self.conn.cursor()
        cur.execute("SELECT S.*, C.name as client_name, C.phone, C.address, P.name as product_name FROM Sales S LEFT JOIN Clients C ON S.client_id=C.client_id LEFT JOIN Products P ON S.product_id=P.product_id WHERE S.sale_id=?", (sale_id,))
        row = cur.fetchone(); 
        if not row: messagebox.showerror("Error", "Sale not found"); return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")])
        if not path: return
        c = canvas.Canvas(path, pagesize=A4); x=50; y=800; c.setFont("Helvetica-Bold",14); c.drawString(x,y,"DZAIR - Invoice"); y-=30; c.setFont("Helvetica",10)
        c.drawString(x,y,f"Invoice: {row['invoice_no']}"); y-=15; c.drawString(x,y,f"Date: {row['date']}"); y-=15
        c.drawString(x,y,f"Client: {row['client_name']} | Phone: {row['phone']}"); y-=15; c.drawString(x,y,f"Address: {row['address']}"); y-=20
        c.drawString(x,y,f"Product: {row['product_name']} x {row['quantity']}"); y-=15; c.drawString(x,y,f"Selling Price: {row['selling_price']}"); y-=15
        c.drawString(x,y,f"Delivery: {row['delivery_price']}"); y-=15; c.drawString(x,y,f"P FAYDA: {row['p_fayda']}"); y-=15; c.drawString(x,y,f"FAYDA SAFIA: {row['fayda_safia']}"); c.showPage(); c.save()
        messagebox.showinfo("Saved", f"Invoice saved to {path}")

    def backup_db(self):
        src = DB_PATH
        dst = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite DB","*.db")])
        if not dst: return
        import shutil; shutil.copy2(src, dst); messagebox.showinfo("Backup", f"DB copied to {dst}")

    def refresh_all(self):
        self.refresh_clients(); self.refresh_products(); self.refresh_sales(); self.refresh_dashboard(); self.refresh_fees()

# ---- run ----
def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
