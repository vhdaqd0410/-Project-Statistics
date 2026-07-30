# -*- coding: utf-8 -*-
"""AI后期剪辑提成表生成工具 - GUI v6.0"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess, threading, os, sys, json, re
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CLI_SCRIPT = os.path.join(SCRIPT_DIR, 'generate_commission.py')
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')
BACKUP_DIR = os.path.join(SCRIPT_DIR, 'backup')
CARDS_DIR = os.path.join(SCRIPT_DIR, '个人绩效卡片')

# 导入功能模块
try:
    from features import (generate_next_month_template, export_to_pdf,
                          generate_person_cards, compare_months, backup_output,
                          data_preview, generate_ranking_html,
                          generate_project_management_html, smart_episode_assignment,
                          generate_person_trend_html, validate_project_data,
                          list_backups, cleanup_backups)
    HAS_FEATURES = True
except ImportError as e:
    HAS_FEATURES = False
    print(f'features import error: {e}')

# 优先使用 Python 3.13（保证有 pandas/openpyxl）
_PY313 = r'C:\Users\Admin\AppData\Local\Programs\Python\Python313\python.exe'
if os.path.exists(_PY313):
    PYTHON_EXE = _PY313
else:
    PYTHON_EXE = sys.executable

ROLES = ['一卡剪辑', '二卡剪辑', '剪辑助理', '剪辑组长', '小组长']
ROLE_ICONS = {'一卡剪辑': '🟢', '二卡剪辑': '🔵', '剪辑助理': '🟣', '剪辑组长': '🟠', '小组长': '🟡'}

# ============ 配色 ============
C = {'bg': '#f5f6fa', 'card': '#ffffff', 'accent': '#4a6cf7',
     'green': '#10b981', 'blue': '#3b82f6', 'purple': '#8b5cf6',
     'orange': '#f59e0b', 'red': '#ef4444', 'gray': '#6b7280',
     'text': '#1f2937', 'sub': '#9ca3af', 'border': '#e5e7eb',
     'hdr_start': '#1e3a5f', 'hdr_end': '#3b82f6',
     'log_bg': '#1a1a2e', 'log_fg': '#e4e4e7', 'btn_hover': '#059669'}

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("AI后期剪辑 提成表生成工具")
        self.root.geometry("900x700")
        self.root.minsize(780, 580)
        self.root.configure(bg=C['bg'])

        # 设置 ttk 样式
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Green.TButton', font=('Microsoft YaHei', 12, 'bold'),
                        background=C['green'], foreground='white',
                        borderwidth=0, padding=(24, 8))
        style.map('Green.TButton', background=[('active', C['btn_hover'])])

        try: self.root.iconbitmap(os.path.join(SCRIPT_DIR, 'icon.ico'))
        except: pass

        self.cfg = self._load_config()
        self.project_file = os.path.join(SCRIPT_DIR, '一组AI项目.xlsx')
        self.template_file = os.path.join(SCRIPT_DIR, 'AI后期剪辑提成一组最新.xlsx')
        self.output_dir = SCRIPT_DIR
        if not os.path.exists(self.template_file):
            self.template_file = os.path.join(SCRIPT_DIR, 'AI后期剪辑提成一组模板.xlsx')
        self.auto_backup = tk.BooleanVar(value=True)
        self._watcher = False
        self.build_ui()

        # 快捷键绑定
        self.root.bind('<Control-g>', lambda e: self.run())
        self.root.bind('<Control-o>', lambda e: os.startfile(self.output_dir))
        self.root.bind('<Control-r>', lambda e: self.open_role_editor())
        self.root.bind('<Control-G>', lambda e: self.run())
        self.root.bind('<Control-O>', lambda e: os.startfile(self.output_dir))
        self.root.bind('<Control-R>', lambda e: self.open_role_editor())
        self.check_files()

    def _load_config(self):
        try:
            with open(CONFIG_PATH, encoding='utf-8') as f: return json.load(f)
        except: return None

    def _save_config(self):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(self.cfg, f, ensure_ascii=False, indent=2)

    def _load_gc_module(self):
        """加载 generate_commission 模块并注入当前配置"""
        import sys as _sys
        _sys.path.insert(0, SCRIPT_DIR)
        import generate_commission as gc
        gc.cfg = self.cfg
        gc.ROLE_MAP = self.cfg['人员角色']
        gc.RULES = self.cfg['rules']
        gc.GROUPS = self.cfg.get('小组', {})
        gc.ALL_NAMES = list(gc.ROLE_MAP.keys())
        gc.NAME_ORDER = self.cfg['人员排序']
        return gc, _sys

    # ============ UI ============

    def build_ui(self):
        hdr = tk.Frame(self.root, bg=C['hdr_start'], height=72)
        hdr.pack(fill='x'); hdr.pack_propagate(False)
        tk.Label(hdr, text='AI后期剪辑提成表生成工具', font=('Microsoft YaHei', 18, 'bold'),
                 fg='white', bg=C['hdr_start']).pack(pady=(10, 0))
        tk.Label(hdr, text='智能计算 · 一键生成 · 可视化仪表盘', font=('Microsoft YaHei', 8),
                 fg='#93c5fd', bg=C['hdr_start']).pack()

        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True, padx=10, pady=(6, 0))

        tab1 = tk.Frame(nb, bg=C['bg']); nb.add(tab1, text='  主面板  ')
        self._build_tab_main(tab1)
        tab2 = tk.Frame(nb, bg=C['bg']); nb.add(tab2, text='  工具箱  ')
        self._build_tab_tools(tab2)
        tab3 = tk.Frame(nb, bg=C['bg']); nb.add(tab3, text='  高级  ')
        self._build_tab_advanced(tab3)

        bar = tk.Frame(self.root, bg='#f1f5f9', height=26)
        bar.pack(fill='x', side='bottom'); bar.pack_propagate(False)
        self.st = tk.Label(bar, text='● 就绪', font=('Microsoft YaHei', 9),
                           bg='#f1f5f9', fg=C['sub'], anchor='w', padx=12)
        self.st.pack(fill='x')

    def _build_tab_main(self, p):
        c1 = tk.Frame(p, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c1.pack(fill='x', padx=10, pady=(8, 5))
        inner = tk.Frame(c1, bg=C['card']); inner.pack(fill='x', padx=12, pady=8)
        tk.Label(inner, text='📋 文件配置', font=('Microsoft YaHei', 11, 'bold'),
                 bg=C['card'], fg=C['text']).pack(anchor='w')
        for label, attr, cmd, is_dir in [
            ('项目数据:', 'pf_label', self._select_project, False),
            ('模板文件:', 'tf_label', self._select_template, False),
            ('输出目录:', 'od_label', self._select_output_dir, True)]:
            row = tk.Frame(inner, bg=C['card']); row.pack(fill='x', pady=2)
            tk.Label(row, text=label, font=('Microsoft YaHei', 9), bg=C['card'],
                     fg=C['text'], width=10, anchor='w').pack(side='left')
            lbl = tk.Label(row, text='', font=('Microsoft YaHei', 8), bg=C['card'], fg=C['sub'], anchor='w')
            lbl.pack(side='left', fill='x', expand=True, padx=4); setattr(self, attr, lbl)
            btn_text = '选择目录' if is_dir else '选择文件'
            tk.Button(row, text=btn_text, font=('Microsoft YaHei', 8),
                      bg=C['purple'] if is_dir else C['blue'], fg='white',
                      relief='flat', cursor='hand2', padx=10, pady=1, command=cmd).pack(side='right')
        self._refresh_file_labels()

        c2 = tk.Frame(p, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c2.pack(fill='x', padx=10, pady=(0, 5))
        in2 = tk.Frame(c2, bg=C['card']); in2.pack(fill='x', padx=12, pady=6)
        tk.Label(in2, text='👥 当前角色分配', font=('Microsoft YaHei', 11, 'bold'),
                 bg=C['card'], fg=C['text']).pack(anchor='w')
        self.role_tags = tk.Frame(in2, bg=C['card']); self.role_tags.pack(fill='x', pady=(4, 0))
        self._refresh_role_tags()

        btn_row = tk.Frame(p, bg=C['bg']); btn_row.pack(fill='x', padx=10, pady=(2, 3))
        self.run_btn = tk.Button(btn_row, text='▶  一键生成提成表', font=('Microsoft YaHei', 12, 'bold'),
                                  bg=C['green'], fg='white', relief='flat', cursor='hand2',
                                  padx=18, pady=7, activebackground=C['btn_hover'], command=self.run)
        self.run_btn.pack(side='left', padx=(0, 6))
        for txt, clr, cmd in [('👥 当前角色分配', C['orange'], self.open_role_editor),
                               ('📂 打开目录', C['blue'], lambda: os.startfile(self.output_dir))]:
            tk.Button(btn_row, text=txt, font=('Microsoft YaHei', 10), bg=clr, fg='white',
                      relief='flat', cursor='hand2', padx=10, pady=7, command=cmd).pack(side='left', padx=(0, 5))
        tk.Checkbutton(btn_row, text='自动备份', variable=self.auto_backup, font=('Microsoft YaHei', 8),
                       bg=C['bg'], fg=C['text'], selectcolor=C['bg']).pack(side='left', padx=(6, 0))
        tk.Label(btn_row, text='Ctrl+G/O/R', font=('Microsoft YaHei', 7),
                 bg=C['bg'], fg=C['sub']).pack(side='right')

        c3 = tk.Frame(p, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c3.pack(fill='both', expand=True, padx=10, pady=(0, 6))
        in3 = tk.Frame(c3, bg=C['card']); in3.pack(fill='both', expand=True, padx=12, pady=6)
        tk.Label(in3, text='📝 运行日志', font=('Microsoft YaHei', 11, 'bold'),
                 bg=C['card'], fg=C['text']).pack(anchor='w')
        lw = tk.Frame(in3, bg=C['log_bg']); lw.pack(fill='both', expand=True, pady=(3, 0))
        self.log_txt = tk.Text(lw, font=('Consolas', 9), bg=C['log_bg'], fg=C['log_fg'],
                               insertbackground='white', relief='flat', padx=8, pady=6,
                               wrap='word', state='disabled')
        self.log_txt.pack(side='left', fill='both', expand=True)
        sb = tk.Scrollbar(lw, command=self.log_txt.yview); sb.pack(side='right', fill='y')
        self.log_txt.configure(yscrollcommand=sb.set)

    def _build_tab_tools(self, p):
        tools = [
            ('📐 智能分集', '#e11d48', self._smart_assign),
            ('📋 数据预览', '#0891b2', self._preview_data),
            ('✅ 数据校验', '#16a34a', self._validate_data),
            ('🔍 项目去重', '#dc2626', self._check_duplicates),
            ('📊 月份对比', C['orange'], self._compare_months),
            ('📤 导出PDF', C['purple'], self._export_pdf),
            ('📅 下月模板', C['accent'], self._gen_next_template),
            ('🏷️ 提成规则', C['gray'], self._edit_rules),
            ('🏆 组内排名', '#d97706', self._gen_ranking),
            ('🗂️ 项目管理', '#7c3aed', self._gen_project_mgmt),
            ('🃏 绩效卡片', C['blue'], self._gen_cards),
            ('🔄 文件监控', '#059669', self._toggle_watch),
        ]
        grid = tk.Frame(p, bg=C['bg']); grid.pack(fill='both', expand=True, padx=12, pady=12)
        for i, (text, color, cmd) in enumerate(tools):
            tk.Button(grid, text=text, font=('Microsoft YaHei', 10), bg=color, fg='white',
                      relief='flat', cursor='hand2', padx=10, pady=8,
                      activebackground='#333', command=cmd).grid(
                      row=i//3, column=i%3, padx=3, pady=3, sticky='ew')
        for c in range(3): grid.grid_columnconfigure(c, weight=1)

    def _build_tab_advanced(self, p):
        c1 = tk.Frame(p, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c1.pack(fill='x', padx=10, pady=(8, 5))
        in1 = tk.Frame(c1, bg=C['card']); in1.pack(fill='x', padx=12, pady=8)
        tk.Label(in1, text='📊 个人月度趋势', font=('Microsoft YaHei', 11, 'bold'),
                 bg=C['card'], fg=C['text']).pack(anchor='w')
        sf = tk.Frame(in1, bg=C['card']); sf.pack(fill='x', pady=(4, 0))
        self._trend_var = tk.StringVar()
        names = sorted(self.cfg.get('人员角色', {}).keys())
        if names:
            self._trend_var.set(names[0])
            tk.OptionMenu(sf, self._trend_var, *names).pack(side='left', padx=(0, 6))
        tk.Button(sf, text='▶  一键生成提成表 Trend', font=('Microsoft YaHei', 9), bg=C['blue'],
                  fg='white', relief='flat', cursor='hand2', padx=10, pady=3,
                  command=self._gen_trend).pack(side='left')

        c2 = tk.Frame(p, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c2.pack(fill='both', expand=True, padx=10, pady=(0, 6))
        in2 = tk.Frame(c2, bg=C['card']); in2.pack(fill='both', expand=True, padx=12, pady=8)
        hf = tk.Frame(in2, bg=C['card']); hf.pack(fill='x')
        tk.Label(hf, text='🗄️ 备份管理', font=('Microsoft YaHei', 11, 'bold'),
                 bg=C['card'], fg=C['text']).pack(side='left')
        tk.Button(hf, text='清理旧备份', font=('Microsoft YaHei', 8), bg=C['red'],
                  fg='white', relief='flat', cursor='hand2', padx=8, pady=1,
                  command=self._manage_backups).pack(side='right')
        self._backup_list = tk.Text(in2, font=('Consolas', 8), bg=C['log_bg'], fg=C['log_fg'],
                                    relief='flat', padx=8, pady=6, height=8, wrap='word')
        self._backup_list.pack(fill='both', expand=True, pady=(3, 0))
        self._refresh_backups()

    def _refresh_role_tags(self):
        for w in self.role_tags.winfo_children():
            w.destroy()
        if not self.cfg: return
        rm = self.cfg.get('人员角色', {})
        if not rm:
            tk.Label(self.role_tags, text='（暂无人员，请点击"角色配置"添加）',
                     font=('Microsoft YaHei', 9), bg=C['card'],
                     fg=C['sub']).pack(anchor='w')
            return

        for role in ROLES:
            names = [n for n, r in rm.items() if r == role]
            if not names:
                continue
            tag = tk.Frame(self.role_tags, bg=C['bg'], highlightthickness=1,
                           highlightbackground=C['border'])
            tag.pack(side='left', padx=(0, 10), pady=4)
            tk.Label(tag, text=f'{ROLE_ICONS[role]} {role} ({len(names)}人)', font=('Microsoft YaHei', 9, 'bold'),
                     bg=C['bg'], fg=C['text']).pack(side='left', padx=(8, 4), pady=4)
            tk.Label(tag, text='、'.join(names), font=('Microsoft YaHei', 9),
                     bg=C['bg'], fg=C['gray']).pack(side='left', padx=(0, 8), pady=4)

    # ============ 角色编辑器 ============

    def open_role_editor(self):
        if not self.cfg:
            messagebox.showwarning('错误', '无法加载 config.json'); return
        dlg = tk.Toplevel(self.root)
        dlg.title('角色配置编辑器'); dlg.geometry('620x620')
        dlg.minsize(500, 500); dlg.configure(bg=C['bg'])
        dlg.transient(self.root); dlg.grab_set()

        tk.Label(dlg, text='👥 人员角色配置', font=('Microsoft YaHei', 16, 'bold'),
                 bg=C['bg'], fg=C['text']).pack(pady=(16, 2))
        tk.Label(dlg, text='修改角色 / 添加人员 / 删除人员 · 保存后即时生效',
                 font=('Microsoft YaHei', 9), bg=C['bg'],
                 fg=C['sub']).pack(pady=(0, 8))

        # ---- 新增人员栏 ----
        add_bar = tk.Frame(dlg, bg=C['bg'])
        add_bar.pack(fill='x', padx=20, pady=(0, 8))
        tk.Label(add_bar, text='新增:', font=('Microsoft YaHei', 10),
                 bg=C['bg'], fg=C['text']).pack(side='left', padx=(0, 6))
        add_name = tk.Entry(add_bar, font=('Microsoft YaHei', 10), width=12)
        add_name.pack(side='left', padx=(0, 6))
        add_role_var = tk.StringVar(value='一卡剪辑')
        ttk.Combobox(add_bar, textvariable=add_role_var, values=ROLES,
                     state='readonly', font=('Microsoft YaHei', 10), width=10).pack(side='left', padx=(0, 6))
        rm_ref = self.cfg['人员角色']
        order_ref = self.cfg.get('人员排序', list(rm_ref.keys()))

        # 刷新函数——重建整个列表
        def rebuild_person_list():
            for w in sf.winfo_children():
                w.destroy()
            for i, name in enumerate(order_ref):
                if name not in rm_ref:
                    continue
                row = tk.Frame(sf, bg='white', highlightthickness=1,
                               highlightbackground=C['border'])
                row.grid(row=i, column=0, padx=5, pady=2, sticky='ew')
                # 删除按钮
                btn_del = tk.Button(row, text='✕', font=('Microsoft YaHei', 9, 'bold'),
                                    bg='#fee2e2', fg='#dc2626', relief='flat',
                                    cursor='hand2', padx=6, pady=3,
                                    activebackground='#fecaca',
                                    command=lambda n=name: _delete_person(n))
                btn_del.pack(side='left', padx=(6, 4), pady=5)
                # 姓名
                tk.Label(row, text=name, font=('Microsoft YaHei', 10, 'bold'),
                         bg='white', fg=C['text'], width=8, anchor='w').pack(side='left', padx=(2, 6), pady=5)
                # 角色下拉
                current = rm_ref.get(name, '一卡剪辑')
                var = tk.StringVar(value=current)
                role_vars[name] = var
                cb = ttk.Combobox(row, textvariable=var, values=ROLES,
                                  state='readonly', font=('Microsoft YaHei', 10), width=10)
                cb.pack(side='left', padx=(0, 8), pady=5)
                # 颜色点
                color_map = {'一卡剪辑': '#27ae60', '二卡剪辑': '#3b82f6',
                             '剪辑助理': '#8b5cf6', '剪辑组长': '#f59e0b'}
                dot = tk.Label(row, text='●', font=('Microsoft YaHei', 14),
                               fg=color_map.get(current, '#333'), bg='white')
                dot.pack(side='left', pady=5)
                def _on_change(*_, lbl=dot, v=var):
                    lbl.configure(fg=color_map.get(v.get(), '#333'))
                var.trace_add('write', _on_change)
            sf.grid_columnconfigure(0, weight=1)

        def _delete_person(name):
            if not messagebox.askyesno('确认删除', f'确定要删除人员 "{name}" 吗？\n\n此操作不可恢复。'):
                return
            rm_ref.pop(name, None)
            if name in order_ref:
                order_ref.remove(name)
            role_vars.pop(name, None)
            self.cfg['人员排序'] = order_ref
            rebuild_person_list()

        def _add_person():
            name = add_name.get().strip()
            if not name:
                messagebox.showwarning('提示', '请输入人员姓名'); return
            if name in rm_ref:
                messagebox.showwarning('提示', f'人员 "{name}" 已存在'); return
            rm_ref[name] = add_role_var.get()
            order_ref.append(name)
            self.cfg['人员排序'] = order_ref
            add_name.delete(0, 'end')
            rebuild_person_list()
            # 滚到底部
            canvas.yview_moveto(1.0)

        tk.Button(add_bar, text='＋ 添加', font=('Microsoft YaHei', 10, 'bold'),
                  bg=C['blue'], fg='white', relief='flat', cursor='hand2',
                  padx=14, pady=3, command=_add_person).pack(side='left', padx=(4, 0))
        # 绑定回车
        add_name.bind('<Return>', lambda e: _add_person())

        # ---- 可滚动人员列表 ----
        canvas = tk.Canvas(dlg, bg=C['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(dlg, orient='vertical', command=canvas.yview)
        sf = tk.Frame(canvas, bg=C['bg'])
        sf.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((10, 0), window=sf, anchor='nw', width=570)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True, padx=(20, 0), pady=(0, 14))
        scrollbar.pack(side='right', fill='y', pady=(0, 14))

        role_vars = {}
        rebuild_person_list()

        # ---- 底部按钮 ----
        bf = tk.Frame(dlg, bg=C['bg'])
        bf.pack(fill='x', padx=20, pady=(0, 16))

        def save():
            for n, v in role_vars.items():
                if n in rm_ref:
                    rm_ref[n] = v.get()
            self.cfg['人员角色'] = rm_ref
            self.cfg['人员排序'] = order_ref

            # 同步更新小组中的成员列表
            all_names = set(rm_ref.keys())
            groups = self.cfg.get('小组', {})
            for gname, ginfo in groups.items():
                members = ginfo.get('成员', [])
                new_members = [m for m in members if m in all_names]
                ginfo['成员'] = new_members
                leader = ginfo.get('组长', '')
                if leader not in all_names:
                    ginfo['组长'] = new_members[0] if new_members else ''

            self._save_config()
            self._refresh_role_tags()
            self._log('✅ 角色配置已更新并保存')
            dlg.destroy()
            messagebox.showinfo('保存成功', '角色配置已保存！\n\n下次生成时生效。')

        tk.Button(bf, text='💾 保存配置', font=('Microsoft YaHei', 12, 'bold'),
                  bg=C['green'], fg='white', relief='flat', cursor='hand2',
                  padx=30, pady=8, command=save).pack(side='right', padx=(10, 0))
        tk.Button(bf, text='取消', font=('Microsoft YaHei', 11),
                  bg=C['gray'], fg='white', relief='flat', cursor='hand2',
                  padx=24, pady=8, command=dlg.destroy).pack(side='right')
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-e.delta/120), 'units'))
        dlg.protocol('WM_DELETE_WINDOW', lambda: [canvas.unbind_all('<MouseWheel>'), dlg.destroy()])

    # ============ 文件选择 ============

    def _select_project(self):
        path = filedialog.askopenfilename(
            title='选择项目数据文件',
            filetypes=[('Excel文件', '*.xlsx'), ('所有文件', '*.*')],
            initialdir=SCRIPT_DIR)
        if path:
            self.project_file = path
            self._refresh_file_labels()
            self.check_files()

    def _select_template(self):
        path = filedialog.askopenfilename(
            title='选择模板文件',
            filetypes=[('Excel文件', '*.xlsx'), ('所有文件', '*.*')],
            initialdir=SCRIPT_DIR)
        if path:
            self.template_file = path
            self._refresh_file_labels()
            self.check_files()

    def _select_output_dir(self):
        path = filedialog.askdirectory(
            title='选择输出目录',
            initialdir=self.output_dir)
        if path:
            self.output_dir = path
            self._refresh_file_labels()
            self.check_files()

    def _refresh_file_labels(self):
        def _short(fpath):
            if not fpath: return '（未选择）'
            name = os.path.basename(fpath)
            exists = os.path.exists(fpath)
            icon = '✅' if exists else '❌'
            return f'{icon} {name}'
        self.pf_label.configure(text=_short(self.project_file))
        self.tf_label.configure(text=_short(self.template_file))
        self.od_label.configure(text=f'📁 {self.output_dir}')

    # ============ 功能：生成下月模板 ============
    def _gen_next_template(self):
        if not HAS_FEATURES:
            messagebox.showerror('错误', '功能模块(features.py)未找到')
            return
        try:
            self._log('📅 正在生成下月模板...')
            self._log(f'   源模板: {os.path.basename(self.template_file)}')
            path, msg = generate_next_month_template(self.template_file, self.output_dir)
            if path:
                self._log(f'✅ {msg}')
                self._log(f'   文件: {os.path.basename(path)}')
                try: os.startfile(path)
                except: pass
                messagebox.showinfo('完成', f'{msg}\n\n文件已自动打开。')
            else:
                self._log(f'❌ {msg}')
                messagebox.showerror('失败', msg)
        except Exception as e:
            self._log(f'❌ 生成下月模板失败: {e}')
            messagebox.showerror('失败', f'生成失败:\n{e}')

    # ============ 功能：导出📤 导出PDF ============
    def _export_pdf(self):
        if not HAS_FEATURES:
            messagebox.showerror('错误', '功能模块(features.py)未找到')
            return
        path = filedialog.askopenfilename(
            title='选择要导出的Excel提成表',
            filetypes=[('Excel文件', '*.xlsx')],
            initialdir=SCRIPT_DIR)
        if not path:
            return
        self._log(f'📤 正在导出📤 导出PDF: {os.path.basename(path)}')
        try:
            pdf_path = export_to_pdf(path)
            self._log(f'✅ 📤 导出PDF已生成: {os.path.basename(pdf_path)}')
            try: os.startfile(pdf_path)
            except: pass
            messagebox.showinfo('完成', f'📤 导出PDF导出成功！\n\n📄 {os.path.basename(pdf_path)}')
        except Exception as e:
            self._log(f'❌ 📤 导出PDF导出失败: {e}')
            messagebox.showerror('失败', f'📤 导出PDF导出失败:\n{e}')

    # ============ 功能：个人绩效卡片 ============
    def _gen_cards(self):
        if not HAS_FEATURES:
            messagebox.showerror('错误', '功能模块(features.py)未找到')
            return
        self._log('🃏 正在生成个人绩效卡片...')
        try:
            import pandas as pd
            gc, _sys = self._load_gc_module()

            df = pd.read_excel(self.project_file, header=None)
            records, group_pids = gc.parse_projects(df)
            if not records:
                self._log('⚠️ 无数据，请检查项目文件')
                _sys.path.pop(0)
                return
            cd = gc.compute_commission(records, group_pids)
            card_paths = generate_person_cards(records, cd, CARDS_DIR)
            _sys.path.pop(0)

            if card_paths:
                self._log(f'✅ 已生成 {len(card_paths)-1} 张个人绩效卡片 -> 个人绩效卡片/')
                try: os.startfile(card_paths[0])
                except: pass
            else:
                self._log('⚠️ 未能生成卡片')
        except Exception as e:
            self._log(f'❌ 卡片生成失败: {e}')

    # ============ 功能：多月份对比 ============
    def _compare_months(self):
        if not HAS_FEATURES:
            messagebox.showerror('错误', '功能模块(features.py)未找到')
            return
        f1 = filedialog.askopenfilename(
            title='选择第一个月份的提成表',
            filetypes=[('Excel文件', '*.xlsx')],
            initialdir=SCRIPT_DIR)
        if not f1: return
        f2 = filedialog.askopenfilename(
            title='选择第二个月份的提成表',
            filetypes=[('Excel文件', '*.xlsx')],
            initialdir=os.path.dirname(f1))
        if not f2: return
        self._log(f'📊 正在对比两个月份的提成表...')
        try:
            label1, label2, diffs = compare_months(f1, f2)
            if not diffs:
                self._log(f'✅ 两个月份数据完全一致，无差异。')
                messagebox.showinfo('对比结果', f'{label1} ↔ {label2}\n\n数据完全一致，无差异。')
                return
            self._log(f'{label1} ↔ {label2} 对比: {len(diffs)}人有变化')
            msg = f'{label1}  →  {label2}\n{"="*60}\n'
            msg += f'{"姓名":　<6s} {"上月集数":>6s} {"本月集数":>6s} {"集数变化":>6s} {"上月提成":>8s} {"本月提成":>8s} {"提成变化":>8s}\n'
            for nm, e1, e2, de, c1, c2, dc in diffs:
                de_str = f'+{de}' if de > 0 else str(de)
                dc_str = f'+{dc}' if dc > 0 else str(dc)
                msg += f'{nm:　<6s} {e1:>6d} {e2:>6d} {de_str:>6s} {c1:>8d} {c2:>8d} {dc_str:>8s}\n'
            self._log(msg)
            # 弹窗显示
            dlg = tk.Toplevel(self.root)
            dlg.title('多月份对比')
            dlg.geometry('700x500')
            dlg.configure(bg=C['bg'])
            tk.Label(dlg, text=f'{label1}  ↔  {label2}', font=('Microsoft YaHei', 14, 'bold'),
                     bg=C['bg'], fg=C['text']).pack(pady=10)
            txt = tk.Text(dlg, font=('Consolas', 10), bg=C['log_bg'], fg=C['log_fg'],
                          relief='flat', padx=12, pady=10)
            txt.insert('1.0', msg)
            txt.configure(state='disabled')
            txt.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        except Exception as e:
            self._log(f'❌ 对比失败: {e}')
            messagebox.showerror('失败', f'对比失败:\n{e}')

    # ============ 功能：提成规则面板 ============
    def _edit_rules(self):
        dlg = tk.Toplevel(self.root)
        dlg.title('提成规则设置')
        dlg.geometry('500x420')
        dlg.configure(bg=C['bg'])
        dlg.transient(self.root); dlg.grab_set()

        tk.Label(dlg, text='🏷️ 提成规则配置', font=('Microsoft YaHei', 15, 'bold'),
                 bg=C['bg'], fg=C['text']).pack(pady=(20, 10))
        tk.Label(dlg, text='修改后点击保存，下次生成生效', font=('Microsoft YaHei', 9),
                 bg=C['bg'], fg=C['sub']).pack()

        rules = self.cfg.get('rules', {})
        vars_map = {}
        for role_name in ['一卡剪辑', '二卡剪辑', '剪辑助理']:
            frm = tk.Frame(dlg, bg=C['card'], highlightthickness=1,
                           highlightbackground=C['border'])
            frm.pack(fill='x', padx=20, pady=4)
            tk.Label(frm, text=f'{role_name}', font=('Microsoft YaHei', 11, 'bold'),
                     bg=C['card'], fg=C['text'], width=12, anchor='w').pack(side='left', padx=10, pady=8)

            r = rules.get(role_name, {})
            for key, label in [('基准集数', '基准'), ('超额每集', '超额/集'), ('缺集每集扣', '缺扣/集')]:
                v = tk.IntVar(value=r.get(key, 70 if role_name == '一卡剪辑' else 120))
                vars_map[f'{role_name}|{key}'] = v
                sub = tk.Frame(frm, bg=C['card'])
                sub.pack(side='left', padx=4, pady=4)
                tk.Label(sub, text=label, font=('Microsoft YaHei', 8), bg=C['card'],
                         fg=C['sub']).pack()
                s = tk.Scale(sub, from_=0, to=300 if '集数' in key else 100,
                             orient='horizontal', variable=v, length=70,
                             bg=C['card'], fg=C['text'], highlightthickness=0)
                s.pack()

        # 组长规则
        frm_l = tk.Frame(dlg, bg=C['card'], highlightthickness=1,
                         highlightbackground=C['border'])
        frm_l.pack(fill='x', padx=20, pady=4)
        tk.Label(frm_l, text='剪辑组长', font=('Microsoft YaHei', 11, 'bold'),
                 bg=C['card'], fg=C['text'], width=12, anchor='w').pack(side='left', padx=10, pady=8)
        r_l = rules.get('剪辑组长', {})
        for key, label in [('每集单价', '单价/集'), ('组内每部提成', '每部奖')]:
            v = tk.IntVar(value=r_l.get(key, 20 if '单价' in key else 100))
            vars_map[f'剪辑组长|{key}'] = v
            sub = tk.Frame(frm_l, bg=C['card'])
            sub.pack(side='left', padx=4, pady=4)
            tk.Label(sub, text=label, font=('Microsoft YaHei', 8), bg=C['card'],
                     fg=C['sub']).pack()
            s = tk.Scale(sub, from_=0, to=300, orient='horizontal', variable=v,
                         length=70, bg=C['card'], fg=C['text'], highlightthickness=0)
            s.pack()

        def save_rules():
            for key, v in vars_map.items():
                role, field = key.split('|', 1)
                rules[role][field] = v.get()
            self._save_config()
            self._log('✅ 提成规则已更新并保存')
            dlg.destroy()
            messagebox.showinfo('保存成功', '提成规则已保存！\n\n下次生成时生效。')

        tk.Button(dlg, text='💾 保存规则', font=('Microsoft YaHei', 12, 'bold'),
                  bg=C['green'], fg='white', relief='flat', cursor='hand2',
                  padx=30, pady=8, command=save_rules).pack(pady=16)

    # ============ 功能：数据预览 ============
    def _preview_data(self):
        self._log('📋 正在加载数据预览...')
        try:
            import pandas as pd
            gc, _sys = self._load_gc_module()

            self._log(f'   项目文件: {os.path.basename(self.project_file)}')
            self._log(f'   人员数量: {len(gc.ALL_NAMES)}')
            df = pd.read_excel(self.project_file, header=None)
            self._log(f'   Excel行数: {len(df)}')
            records, group_pids = gc.parse_projects(df)
            self._log(f'   解析记录: {len(records)}')
            cd = gc.compute_commission(records, group_pids)
            preview = data_preview(records, cd)
            self._log(f'   预览行数: {len(preview)}')
            _sys.path.pop(0)

            if not preview:
                self._log('⚠️ 无数据可预览（可能项目文件名不在config人员名单中）')
                messagebox.showwarning('无数据', '解析到0条记录。\n\n请确认项目数据文件中的人员姓名与角色配置中的一致。')
                return

            # 弹窗展示
            dlg = tk.Toplevel(self.root)
            dlg.title('数据预览')
            dlg.geometry('750x520')
            dlg.configure(bg=C['bg'])
            tk.Label(dlg, text='📋 数据预览', font=('Microsoft YaHei', 15, 'bold'),
                     bg=C['bg'], fg=C['text']).pack(pady=12)

            # Treeview
            cols = ('姓名', '角色', '集数', '项目数', '基准', '绩效', '提成')
            tree = ttk.Treeview(dlg, columns=cols, show='headings', height=18)
            widths = [80, 90, 60, 60, 60, 60, 80]
            for c, w in zip(cols, widths):
                tree.heading(c, text=c)
                tree.column(c, width=w, anchor='center')
            tree.pack(fill='both', expand=True, padx=20, pady=(0, 10))

            # 颜色标签
            tree.tag_configure('ok', foreground='#16a34a')
            tree.tag_configure('fail', foreground='#dc2626')
            for p in preview:
                tag = 'ok' if p['status'] == '是' else 'fail'
                tree.insert('', 'end', values=(
                    p['name'], p['role'], p['episodes'], p['projects'],
                    f'{p["quota"]}集' if p['quota'] > 0 else '无', p['status'],
                    f'{p["commission"]:,}'), tags=(tag,))
        except Exception as e:
            self._log(f'❌ 预览失败: {e}')
            import traceback; self._log(traceback.format_exc())

    # ============ 功能：组内排名 ============
    def _gen_ranking(self):
        self._log('🏆 正在生成组内排名...')
        try:
            import pandas as pd
            gc, _sys = self._load_gc_module()

            df = pd.read_excel(self.project_file, header=None)
            records, group_pids = gc.parse_projects(df)
            cd = gc.compute_commission(records, group_pids)
            _sys.path.pop(0)

            path = os.path.join(self.output_dir, '组内排名.html')
            generate_ranking_html(cd, path)
            self._log(f'✅ 组内排名已生成: {os.path.basename(path)}')
            try: os.startfile(path)
            except: pass
        except Exception as e:
            self._log(f'❌ 排名生成失败: {e}')

    # ============ 功能：项目管理视图 ============
    def _gen_project_mgmt(self):
        self._log('🗂️ 正在生成项目管理视图...')
        try:
            import pandas as pd
            gc, _sys = self._load_gc_module()

            df = pd.read_excel(self.project_file, header=None)
            records, _ = gc.parse_projects(df)
            _sys.path.pop(0)

            path = os.path.join(self.output_dir, '项目管理_项目清单.html')
            generate_project_management_html(records, path)
            self._log(f'✅ 项目管理视图已生成: {os.path.basename(path)}')
            try: os.startfile(path)
            except: pass
        except Exception as e:
            self._log(f'❌ 项目管理视图生成失败: {e}')

    # ============ 功能：文件监控 ============
    def _toggle_watch(self):
        if hasattr(self, '_watcher') and self._watcher:
            self._stop_watch()
            return
        self._watcher = True
        self._watch_mtime = os.path.getmtime(self.project_file)
        self.btn_watch.configure(text='🔄 监控中...', bg='#059669')
        self._log('🔄 文件监控已开启，检测到项目文件变化将自动提示...')
        self._watch_loop()

    def _watch_loop(self):
        if not getattr(self, '_watcher', False):
            return
        try:
            current_mtime = os.path.getmtime(self.project_file)
            if current_mtime != self._watch_mtime:
                self._watch_mtime = current_mtime
                self._log('🔔 检测到项目数据文件已更新！')
                if messagebox.askyesno('文件更新', '项目数据文件已更新，是否立即重新生成？'):
                    self.run()
        except Exception:
            pass
        if getattr(self, '_watcher', False):
            self.root.after(5000, self._watch_loop)

    def _stop_watch(self):
        self._watcher = False
        self.btn_watch.configure(text='🔄 开启监控', bg='#059669')
        self._log('🔴 文件监控已关闭')

    # ============ 功能：智能分集 ============
    def _smart_assign(self):
        if not HAS_FEATURES:
            messagebox.showerror('错误', '功能模块(features.py)未找到')
            return

        dlg = tk.Toplevel(self.root)
        dlg.title('📐 智能分集')
        dlg.geometry('750x680')
        dlg.minsize(650, 550)
        dlg.configure(bg=C['bg'])
        dlg.transient(self.root); dlg.grab_set()

        # 标题
        tk.Label(dlg, text='📐 智能分集工具', font=('Microsoft YaHei', 15, 'bold'),
                 bg=C['bg'], fg=C['text']).pack(pady=(10, 1))
        tk.Label(dlg, text='输入项目信息，选择剪辑人员，自动按角色区间分集',
                 font=('Microsoft YaHei', 8), bg=C['bg'], fg=C['sub']).pack()

        # 项目信息卡片
        c1 = tk.Frame(dlg, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c1.pack(fill='x', padx=20, pady=(6, 4))
        in1 = tk.Frame(c1, bg=C['card']); in1.pack(fill='x', padx=12, pady=6)

        tk.Label(in1, text='项目名称:', font=('Microsoft YaHei', 9), bg=C['card'],
                 fg=C['text']).grid(row=0, column=0, sticky='w', pady=2)
        name_var = tk.StringVar()
        tk.Entry(in1, textvariable=name_var, font=('Microsoft YaHei', 10), width=40,
                 relief='solid', borderwidth=1).grid(row=0, column=1, padx=(6, 0), pady=2)

        tk.Label(in1, text='总集数:', font=('Microsoft YaHei', 9), bg=C['card'],
                 fg=C['text']).grid(row=1, column=0, sticky='w', pady=2)
        eps_var = tk.IntVar(value=100)
        tk.Spinbox(in1, from_=1, to=2000, textvariable=eps_var, font=('Microsoft YaHei', 10),
                   width=8, relief='solid', borderwidth=1).grid(row=1, column=1, padx=(6, 0), pady=2, sticky='w')

        tk.Label(in1, text='一卡区间(前N集):', font=('Microsoft YaHei', 9), bg=C['card'],
                 fg=C['text']).grid(row=2, column=0, sticky='w', pady=2)
        range_var = tk.IntVar(value=15)
        tk.Spinbox(in1, from_=1, to=500, textvariable=range_var, font=('Microsoft YaHei', 10),
                   width=8, relief='solid', borderwidth=1).grid(row=2, column=1, padx=(6, 0), pady=2, sticky='w')

        # 人员选择卡片
        c2 = tk.Frame(dlg, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c2.pack(fill='x', padx=20, pady=(2, 2))
        in2 = tk.Frame(c2, bg=C['card']); in2.pack(fill='x', padx=12, pady=6)
        tk.Label(in2, text='选择剪辑人员:', font=('Microsoft YaHei', 9, 'bold'),
                 bg=C['card'], fg=C['text']).pack(anchor='w')

        cb_frame = tk.Frame(in2, bg=C['card'])
        cb_frame.pack(fill='x', pady=(3, 0))

        check_vars = {}
        roles_map = self.cfg.get('人员角色', {})
        # 按角色排序展示
        role_order = {'小组长': 0, '一卡剪辑': 1, '二卡剪辑': 2, '剪辑助理': 3, '剪辑组长': 4}
        sorted_people = sorted(roles_map.keys(), key=lambda n: role_order.get(roles_map[n], 99))

        # 分组显示
        role_labels = [
            ('小组长', '🟡', lambda r: '小组长' in r),
            ('一卡剪辑', '🟢', lambda r: '一卡' in r and '小组长' not in r),
            ('二卡剪辑', '🔵', lambda r: '二卡' in r),
            ('剪辑助理', '🟣', lambda r: '助理' in r),
            ('剪辑组长', '🟠', lambda r: '组长' in r and '小组长' not in r),
        ]
        for role_name, icon, matcher in role_labels:
            people_in_role = [n for n in sorted_people if matcher(roles_map[n])]
            if not people_in_role: continue
            tk.Label(cb_frame, text=f'{icon} {role_name}', font=('Microsoft YaHei', 8),
                           bg=C['card'], fg=C['text']).pack(anchor='w', pady=(3, 0))
            sub_frame = tk.Frame(cb_frame, bg=C['card'])
            sub_frame.pack(fill='x')
            for nm in people_in_role:
                v = tk.BooleanVar(value=False)
                check_vars[nm] = v
                tk.Checkbutton(sub_frame, text=nm, variable=v, font=('Microsoft YaHei', 8),
                               bg=C['card'], fg=C['text'], selectcolor=C['card'],
                               activebackground=C['card']).pack(side='left', padx=(0, 8))

        # 全选/全不选按钮
        sel_frame = tk.Frame(in2, bg=C['card'])
        sel_frame.pack(fill='x', pady=(4, 0))
        def _select_all():
            for v in check_vars.values(): v.set(True)
        def _select_none():
            for v in check_vars.values(): v.set(False)
        def _select_card1():
            for n, v in check_vars.items():
                r = roles_map.get(n, '')
                v.set('一卡' in r or '小组长' in r)
        def _select_card2():
            for n, v in check_vars.items():
                r = roles_map.get(n, '')
                v.set('二卡' in r or '助理' in r or ('组长' in r and '小组长' not in r))

        tk.Button(sel_frame, text='全选', font=('Microsoft YaHei', 8),
                  bg=C['blue'], fg='white', relief='flat', cursor='hand2',
                  padx=10, pady=2, command=_select_all).pack(side='left', padx=(0, 6))
        tk.Button(sel_frame, text='全不选', font=('Microsoft YaHei', 8),
                  bg=C['gray'], fg='white', relief='flat', cursor='hand2',
                  padx=10, pady=2, command=_select_none).pack(side='left', padx=(0, 6))
        tk.Button(sel_frame, text='小组长+一卡', font=('Microsoft YaHei', 8),
                  bg='#eab308', fg='white', relief='flat', cursor='hand2',
                  padx=10, pady=2, command=_select_card1).pack(side='left', padx=(0, 6))
        tk.Button(sel_frame, text='仅二卡/助理/组长', font=('Microsoft YaHei', 8),
                  bg='#d97706', fg='white', relief='flat', cursor='hand2',
                  padx=10, pady=2, command=_select_card2).pack(side='left')

        # 结果展示区
        c3 = tk.Frame(dlg, bg=C['card'], highlightthickness=1, highlightbackground=C['border'])
        c3.pack(fill='both', expand=True, padx=20, pady=(4, 12))
        in3 = tk.Frame(c3, bg=C['card']); in3.pack(fill='both', expand=True, padx=12, pady=6)
        tk.Label(in3, text='📋 分集结果:', font=('Microsoft YaHei', 9, 'bold'),
                 bg=C['card'], fg=C['text']).pack(anchor='w')

        result_text = tk.Text(in3, font=('Consolas', 9), bg=C['log_bg'], fg=C['log_fg'],
                              relief='flat', padx=8, pady=6, height=6, wrap='word')
        result_text.pack(fill='both', expand=True, pady=(2, 0))

        # 执行函数
        last_result = [None]  # mutable container

        def _display_result(result, project_name):
            result_text.delete('1.0', 'end')
            result_text.insert('end', f'项目: {project_name}  ({eps_var.get()}集, 一卡区间1～{range_var.get()}集)\n')
            result_text.insert('end', '=' * 60 + '\n\n')
            result_text.insert('end', '📊 统计:\n')
            for k, v in result['stats'].items():
                result_text.insert('end', f'   {k}: {v}\n')
            result_text.insert('end', '\n📋 分集明细 (可直接复制到项目Excel):\n')
            result_text.insert('end', '-' * 60 + '\n')

            selected = [n for n, v in check_vars.items() if v.get()]
            role_labels2 = [
                ('小组长', lambda r: '小组长' in r),
                ('一卡剪辑', lambda r: '一卡' in r and '小组长' not in r),
                ('二卡剪辑', lambda r: '二卡' in r),
                ('剪辑助理', lambda r: '助理' in r),
                ('剪辑组长', lambda r: '组长' in r and '小组长' not in r),
            ]
            for role_name, matcher in role_labels2:
                people = [n for n in selected if matcher(roles_map.get(n, ''))]
                if not people: continue
                result_text.insert('end', f'\n【{role_name}】\n')
                for nm in people:
                    fmt = result['formatted'].get(nm, '')
                    cnt = result['summary'].get(nm, 0)
                    result_text.insert('end', f'  {nm}（{cnt}集）: {fmt}\n')

            result_text.insert('end', '\n\n📝 完整粘贴行 (直接复制到项目Excel分配列):\n')
            result_text.insert('end', '-' * 60 + '\n')
            lines = []
            for nm in selected:
                fmt = result['formatted'].get(nm, '')
                if fmt:
                    lines.append(f'{nm}: {fmt}')
            result_text.insert('end', '\n'.join(lines))

        def _save_txt(proj_name, result):
            """保存分集结果为TXT"""
            try:
                safe_name = re.sub(r'[\\/:*?"<>|]', '_', proj_name)
                txt_path = os.path.join(self.output_dir, f'{safe_name}_分集.txt')
                lines = []
                lines.append(f'项目: {proj_name}  ({eps_var.get()}集, 一卡区间1～{range_var.get()}集)')
                lines.append('=' * 60)
                lines.append('')
                for rn, matcher in [('小组长', lambda r: '小组长' in r),
                                     ('一卡剪辑', lambda r: '一卡' in r and '小组长' not in r),
                                     ('二卡剪辑', lambda r: '二卡' in r),
                                     ('剪辑助理', lambda r: '助理' in r),
                                     ('剪辑组长', lambda r: '组长' in r and '小组长' not in r)]:
                    people = [n for n in result['formatted'] if matcher(roles_map.get(n, ''))]
                    if not people: continue
                    lines.append(f'【{rn}】')
                    for nm in people:
                        lines.append(f'  {nm}（{result["summary"][nm]}集）: {result["formatted"][nm]}')
                    lines.append('')
                lines.append('--- 粘贴行 ---')
                for nm in result['formatted']:
                    f = result['formatted'].get(nm, '')
                    if f: lines.append(f'{nm}: {f}')
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(lines))
                self._log(f'📄 分集TXT已保存: {os.path.basename(txt_path)}')
            except Exception as e:
                self._log(f'⚠️ 保存TXT失败: {e}')

        def _append_to_project(proj_name, result, selected_people, total_eps_val):
            """将分集结果追加到项目数据Excel，格式与现有数据一致"""
            try:
                from openpyxl import load_workbook
                from openpyxl.styles import Font, Alignment
                from datetime import datetime, timedelta

                now = datetime.now()
                end_date = now + timedelta(days=1)
                date_str = f'{end_date.month}.{end_date.day}下午18点交'
                # 项目名不带AI前缀，仿照现有格式
                proj_label = proj_name
                dir_label = f'O:\\AI漫剧剪辑一组\\{proj_name}'

                font_title = Font(name='宋体', size=14, bold=True)
                font_normal = Font(name='宋体', size=14)
                align_center = Alignment(horizontal='center', vertical='center')

                wb = load_workbook(self.project_file)
                ws = wb.active

                # 找到末尾，空两行后写入
                last = ws.max_row
                for r in range(last, 0, -1):
                    if ws.cell(r, 1).value or ws.cell(r, 3).value:
                        last = r
                        break
                next_row = last + 2  # 空一行隔开
                title_row = next_row

                # --- 项目标题行 ---
                c = ws.cell(next_row, 1, proj_label)
                c.font = font_title; c.alignment = align_center

                c = ws.cell(next_row, 2, dir_label)
                c.font = font_title; c.alignment = align_center

                c = ws.cell(next_row, 4, date_str)
                c.font = font_title; c.alignment = align_center

                c = ws.cell(next_row, 5, '已分集')
                c.font = font_title; c.alignment = align_center
                next_row += 1

                # --- 人员分配行 ---
                data_start = next_row
                for nm in selected_people:
                    fmt = result['formatted'].get(nm, '')
                    if not fmt: continue
                    c = ws.cell(next_row, 3, f'{nm}：{fmt}')
                    c.font = font_normal; c.alignment = align_center
                    next_row += 1
                data_end = next_row - 1

                # --- 合并单元格 (A/B/D/E列跨所有人员行) ---
                if data_end >= title_row:
                    for col in ['A', 'B', 'D', 'E']:
                        ws.merge_cells(f'{col}{title_row}:{col}{data_end}')

                wb.save(self.project_file)
                wb.close()

                self._log(f'📎 已追加到项目数据: {os.path.basename(self.project_file)}')
                try: os.startfile(self.project_file)
                except: pass
            except Exception as e:
                self._log(f'⚠️ 追加项目数据失败: {e}')

        def _do_assign_common():
            """通用分集逻辑 - do_assign 和 reroll 共用"""
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning('提示', '请输入项目名称'); return
            selected = [n for n, v in check_vars.items() if v.get()]
            if not selected:
                messagebox.showwarning('提示', '请至少选择一位剪辑人员'); return
            total = eps_var.get()
            rng = range_var.get()
            result = smart_episode_assignment(total, selected, roles_map, rng)
            last_result[0] = result
            _display_result(result, name)
            _save_txt(name, result)
            _append_to_project(name, result, selected, total)

        def copy_result():
            text = result_text.get('1.0', 'end-1c')
            if text.strip():
                dlg.clipboard_clear()
                dlg.clipboard_append(text)
                messagebox.showinfo('已复制', '分集结果已复制到剪贴板，可直接粘贴到项目Excel。')

        btn_f = tk.Frame(dlg, bg=C['bg'])
        btn_f.pack(fill='x', padx=20, pady=(0, 14))
        tk.Button(btn_f, text='🎲 随机分集', font=('Microsoft YaHei', 12, 'bold'),
                  bg='#e11d48', fg='white', relief='flat', cursor='hand2',
                  padx=24, pady=8, activebackground='#be123c',
                  command=_do_assign_common).pack(side='left', padx=(0, 10))
        tk.Button(btn_f, text='� 再次随机', font=('Microsoft YaHei', 11),
                  bg='#d97706', fg='white', relief='flat', cursor='hand2',
                  padx=16, pady=8, activebackground='#b45309',
                  command=_do_assign_common).pack(side='left', padx=(0, 10))
        tk.Button(btn_f, text='📋 复制结果', font=('Microsoft YaHei', 11),
                  bg=C['blue'], fg='white', relief='flat', cursor='hand2',
                  padx=16, pady=8, command=copy_result).pack(side='left', padx=(0, 10))
        tk.Button(btn_f, text='关闭', font=('Microsoft YaHei', 11),
                  bg=C['gray'], fg='white', relief='flat', cursor='hand2',
                  padx=16, pady=8, command=dlg.destroy).pack(side='right')

    def _check_duplicates(self):
        self._log('🔍 正在检查项目去重...')
        try:
            import pandas as pd
            gc, _sys = self._load_gc_module()

            df = pd.read_excel(self.project_file, header=None)
            records, _ = gc.parse_projects(df)
            if not records:
                self._log('⚠️ 未解析到有效记录')
                _sys.path.pop(0)
                return

            ok = gc.self_check(records)
            id_names = {}
            for r in records:
                pid = r['项目ID']
                if pid:
                    if pid not in id_names:
                        id_names[pid] = set()
                    id_names[pid].add(r['AI项目名称'][:30])

            issues = {pid: names for pid, names in id_names.items() if len(names) > 1}
            if issues:
                msg = f'⚠️ 发现 {len(issues)} 个项目ID存在名称不一致:\n'
                for pid, names in list(issues.items())[:10]:
                    msg += f'  ID={pid}: {names}\n'
                self._log(msg)
                messagebox.showwarning('去重检查', msg)
            else:
                self._log(f'✅ 全部 {len(id_names)} 个项目ID名称一致，无重复。')
                messagebox.showinfo('去重检查', f'✅ 检查通过！\n\n{len(id_names)} 个项目ID，名称均一致，无重复。')
            _sys.path.pop(0)
        except Exception as e:
            self._log(f'❌ 检查失败: {e}')
            messagebox.showerror('失败', str(e))

    # ============ 新增功能：数据校验 ============
    def _validate_data(self):
        self._log('🔍 正在校验项目数据...')
        try:
            issues = validate_project_data(self.project_file, self.cfg.get('人员角色', {}))
            if not issues:
                self._log('✅ 数据校验通过，无问题！')
                messagebox.showinfo('校验通过', '✅ 项目数据校验通过！\n\n未发现日期异常、人名不匹配、或集数重复分配等问题。')
                return

            self._log(f'⚠️ 发现 {len(issues)} 个问题:')
            for loc, item, detail in issues[:20]:
                self._log(f'  [{loc}] {item}: {detail}')

            dlg = tk.Toplevel(self.root)
            dlg.title('数据校验结果')
            dlg.geometry('650x400')
            dlg.configure(bg=C['bg'])
            tk.Label(dlg, text=f'⚠️ 发现 {len(issues)} 个问题', font=('Microsoft YaHei', 14, 'bold'),
                     bg=C['bg'], fg=C['text']).pack(pady=12)
            txt = tk.Text(dlg, font=('Consolas', 10), bg=C['log_bg'], fg=C['log_fg'],
                          relief='flat', padx=10, pady=8)
            for loc, item, detail in issues[:50]:
                txt.insert('end', f'[{loc}] {item}: {detail}\n')
            txt.configure(state='disabled')
            txt.pack(fill='both', expand=True, padx=20, pady=(0, 12))
        except Exception as e:
            self._log(f'❌ 校验失败: {e}')

    # ============ 新增功能：个人趋势 ============
    def _gen_trend(self):
        name = self._trend_var.get()
        if not name:
            messagebox.showwarning('提示', '请选择人员'); return
        self._log(f'📊 正在生成 {name} 的月度趋势...')
        try:
            roles_map = self.cfg.get('人员角色', {})
            path = generate_person_trend_html(name, roles_map, self.output_dir)
            if path:
                self._log(f'✅ 趋势图已生成: {os.path.basename(path)}')
                try: os.startfile(path)
                except: pass
            else:
                self._log('⚠️ 未找到该人员的历史数据（可能需要先生成几个月的提成表）')
                messagebox.showinfo('提示', '未找到历史数据。\n\n请确保在输出目录中有多个以 "AI后期剪辑提成一组" 开头的 Excel 文件。')
        except Exception as e:
            self._log(f'❌ 趋势生成失败: {e}')

    # ============ 新增功能：备份管理 ============
    def _refresh_backups(self):
        if not HAS_FEATURES: return
        backups = list_backups(BACKUP_DIR)
        self._backup_list.configure(state='normal')
        self._backup_list.delete('1.0', 'end')
        if not backups:
            self._backup_list.insert('1.0', '暂无备份文件')
        else:
            total_size = sum(b['size'] for b in backups)
            self._backup_list.insert('1.0', f'共 {len(backups)} 个备份文件，占用 {total_size/1024:.1f} KB\n')
            self._backup_list.insert('end', '-' * 50 + '\n')
            for b in backups[:30]:
                self._backup_list.insert('end',
                    f'{b["mtime"].strftime("%m/%d %H:%M")}  {b["size"]/1024:6.1f}KB  {b["name"]}\n')
        self._backup_list.configure(state='disabled')

    def _manage_backups(self):
        if not HAS_FEATURES: return
        removed = cleanup_backups(BACKUP_DIR, keep=30)
        self._log(f'🗄️ 已清理 {removed} 个旧备份，保留最近30个')
        self._refresh_backups()
        messagebox.showinfo('完成', f'已清理 {removed} 个旧备份文件。\n当前保留最近30个。')

    def check_files(self):
        items = [
            (self.project_file, '项目数据'),
            (CONFIG_PATH, '角色配置'),
            (self.template_file, '模板文件'),
        ]
        ok = True
        for fpath, desc in items:
            if not os.path.exists(fpath):
                ok = False
        self.run_btn.configure(state='normal' if ok else 'disabled',
                                bg=C['green'] if ok else C['gray'])
        if ok:
            self._log('✅ 文件就绪，点击"一键生成"开始。')
        else:
            self._log('❌ 文件不完整，请检查。')

    # ============ 日志 ============

    def _log(self, msg):
        self.log_txt.configure(state='normal')
        self.log_txt.insert('end', f'[{datetime.now().strftime("%H:%M:%S")}] {msg}\n')
        self.log_txt.see('end'); self.log_txt.configure(state='disabled')

    def _filter_line(self, line):
        line = line.strip()
        if not line or line.startswith('PS '):
            return
        skip = ['所在位置', 'CategoryInfo', 'FullyQualifiedErrorId', 'NativeCommandError',
                'RemoteException', '~~~~~~~~~~', '鎸変换鎰忛', 'EOFError']
        if any(s in line for s in skip): return
        self._log(line)

    # ============ 生成 ============

    def run(self):
        self.run_btn.configure(state='disabled', bg=C['gray'], text='⏳ 生成中...')
        self.st.configure(text='⏳ 正在计算绩效和生成表格...')
        self._log('🔍 开始执行生成流程...')
        self._log('—' * 50)

        self._output_files = {}  # 收集输出文件路径

        def worker():
            try:
                # 强制 UTF-8 输出，避免 GBK 编码错误（emoji 等字符）
                env = os.environ.copy()
                env['PYTHONIOENCODING'] = 'utf-8'
                p = subprocess.Popen([PYTHON_EXE, CLI_SCRIPT,
                                      self.project_file, self.template_file,
                                      self.output_dir],
                                     cwd=SCRIPT_DIR,
                                     stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                     stderr=subprocess.STDOUT, text=True,
                                     encoding='utf-8', errors='replace',
                                     env=env)
                # 输入回车防止 input() 阻塞
                try: p.stdin.write('\n'); p.stdin.flush()
                except: pass
                for ln in iter(p.stdout.readline, ''):
                    line = ln.strip()
                    # 捕获输出文件路径
                    if line.startswith('OUTPUT_EXCEL='):
                        self._output_files['excel'] = line.split('=', 1)[1]
                    elif line.startswith('OUTPUT_HTML='):
                        self._output_files['html'] = line.split('=', 1)[1]
                    self.root.after(0, lambda l=ln: self._filter_line(l))
                p.wait()
                self.root.after(0, self._done)
            except Exception as e:
                self.root.after(0, lambda: self._log(f'❌ 错误: {e}'))
                self.root.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.run_btn.configure(state='normal', bg=C['green'], text='▶  一键生成提成表')
        self._log('—' * 50)

        excel_path = self._output_files.get('excel', '')
        html_path = self._output_files.get('html', '')

        excel_ok = excel_path and os.path.exists(excel_path)
        html_ok = html_path and os.path.exists(html_path)

        if excel_ok or html_ok:
            self._log('🎉 全部完成！')
            self.st.configure(text='● 就绪 · 生成完成')

            # 自动备份
            if self.auto_backup.get() and HAS_FEATURES:
                try:
                    backed = backup_output(excel_path, html_path, BACKUP_DIR)
                    if backed:
                        self._log(f'💾 已备份 {len(backed)} 个文件到 backup/')
                except Exception as e:
                    self._log(f'⚠️ 备份失败: {e}')

            # 自动生成个人绩效卡片
            if excel_ok and HAS_FEATURES:
                try:
                    import pandas as pd
                    gc, _sys = self._load_gc_module()

                    df = pd.read_excel(self.project_file, header=None)
                    records, group_pids = gc.parse_projects(df)
                    cd = gc.compute_commission(records, group_pids)
                    card_paths = generate_person_cards(records, cd, CARDS_DIR)
                    if card_paths:
                        self._log(f'🃏 已生成 {len(card_paths)-1} 张个人绩效卡片 -> 个人绩效卡片/')
                        try: os.startfile(card_paths[0])
                        except: pass
                    _sys.path.pop(0)
                except Exception as e:
                    self._log(f'⚠️ 卡片生成跳过: {e}')

            # 打开文件
            if excel_ok:
                try: os.startfile(excel_path)
                except: pass
            if html_ok:
                try: os.startfile(html_path)
                except: pass
            messagebox.showinfo('完成', f'✅ 提成表、统计简报和仪表盘已生成完毕！\n\n📊 {os.path.basename(excel_path)}\n📈 {os.path.basename(html_path)}')
        else:
            self._log('❌ 生成失败：未找到输出文件')
            self.st.configure(text='❌ 生成失败 · 请查看日志')
            messagebox.showerror('失败', '❌ 生成失败！\n\n请检查日志了解详细错误信息。\n常见原因：\n  1. 项目数据文件格式不正确\n  2. 模板文件表头不匹配\n  3. config.json 人员配置有误')


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == '__main__':
    main()
