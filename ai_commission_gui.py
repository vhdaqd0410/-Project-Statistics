# -*- coding: utf-8 -*-
"""AI后期剪辑提成表生成工具 - GUI v6.0"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess, threading, os, sys, json
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CLI_SCRIPT = os.path.join(SCRIPT_DIR, 'generate_commission.py')
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')
PYTHON_EXE = sys.executable

ROLES = ['一卡剪辑', '二卡剪辑', '剪辑助理', '剪辑组长']
ROLE_ICONS = {'一卡剪辑': '🟢', '二卡剪辑': '🔵', '剪辑助理': '🟣', '剪辑组长': '🟠'}

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
        if not os.path.exists(self.template_file):
            self.template_file = os.path.join(SCRIPT_DIR, 'AI后期剪辑提成一组模板.xlsx')
        self.build_ui()
        self.check_files()

    def _load_config(self):
        try:
            with open(CONFIG_PATH, encoding='utf-8') as f: return json.load(f)
        except: return None

    def _save_config(self):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(self.cfg, f, ensure_ascii=False, indent=2)

    # ============ UI ============

    def build_ui(self):
        # ---- 顶部渐变标题 ----
        hdr = tk.Frame(self.root, bg=C['hdr_start'], height=100)
        hdr.pack(fill='x'); hdr.pack_propagate(False)

        tk.Label(hdr, text='🎬 AI后期剪辑提成表生成工具',
                 font=('Microsoft YaHei', 22, 'bold'), fg='white',
                 bg=C['hdr_start']).pack(pady=(16, 2))
        tk.Label(hdr, text='智能计算 · 一键生成 · 可视化仪表盘',
                 font=('Microsoft YaHei', 10), fg='#93c5fd',
                 bg=C['hdr_start']).pack()

        # ---- 主体 ----
        main = tk.Frame(self.root, bg=C['bg'])
        main.pack(fill='both', expand=True, padx=24, pady=18)

        # === 卡片1：文件选择 ===
        c1 = tk.Frame(main, bg=C['card'], highlightthickness=1,
                      highlightbackground=C['border'])
        c1.pack(fill='x', pady=(0, 12))

        inner = tk.Frame(c1, bg=C['card'])
        inner.pack(fill='x', padx=20, pady=14)

        tk.Label(inner, text='📋 文件配置',
                 font=('Microsoft YaHei', 13, 'bold'), bg=C['card'],
                 fg=C['text']).pack(anchor='w')

        # 项目文件行
        pf_row = tk.Frame(inner, bg=C['card'])
        pf_row.pack(fill='x', pady=(8, 4))
        tk.Label(pf_row, text='项目数据:', font=('Microsoft YaHei', 10),
                 bg=C['card'], fg=C['text'], width=10, anchor='w').pack(side='left')
        self.pf_label = tk.Label(pf_row, text='', font=('Microsoft YaHei', 9),
                                  bg=C['card'], fg=C['sub'], anchor='w')
        self.pf_label.pack(side='left', fill='x', expand=True, padx=(4, 8))
        tk.Button(pf_row, text='选择文件', font=('Microsoft YaHei', 9),
                  bg=C['blue'], fg='white', relief='flat', cursor='hand2',
                  padx=12, pady=2, activebackground='#2563eb',
                  command=self._select_project).pack(side='right')

        # 模板文件行
        tf_row = tk.Frame(inner, bg=C['card'])
        tf_row.pack(fill='x', pady=(2, 0))
        tk.Label(tf_row, text='模板文件:', font=('Microsoft YaHei', 10),
                 bg=C['card'], fg=C['text'], width=10, anchor='w').pack(side='left')
        self.tf_label = tk.Label(tf_row, text='', font=('Microsoft YaHei', 9),
                                  bg=C['card'], fg=C['sub'], anchor='w')
        self.tf_label.pack(side='left', fill='x', expand=True, padx=(4, 8))
        tk.Button(tf_row, text='选择文件', font=('Microsoft YaHei', 9),
                  bg=C['blue'], fg='white', relief='flat', cursor='hand2',
                  padx=12, pady=2, activebackground='#2563eb',
                  command=self._select_template).pack(side='right')

        self._refresh_file_labels()

        # === 卡片2：角色预览 ===
        c2 = tk.Frame(main, bg=C['card'], highlightthickness=1,
                      highlightbackground=C['border'])
        c2.pack(fill='x', pady=(0, 12))

        inner2 = tk.Frame(c2, bg=C['card'])
        inner2.pack(fill='x', padx=20, pady=14)

        hdr_row = tk.Frame(inner2, bg=C['card'])
        hdr_row.pack(fill='x')
        tk.Label(hdr_row, text='👥 当前角色分配',
                 font=('Microsoft YaHei', 13, 'bold'), bg=C['card'],
                 fg=C['text']).pack(side='left')

        self.role_tags = tk.Frame(inner2, bg=C['card'])
        self.role_tags.pack(fill='x', pady=(10, 0))
        self._refresh_role_tags()

        # === 按钮区 ===
        btn_row = tk.Frame(main, bg=C['bg'])
        btn_row.pack(fill='x', pady=(0, 14))

        self.run_btn = tk.Button(btn_row, text='▶  一键生成提成表',
                                  font=('Microsoft YaHei', 13, 'bold'),
                                  bg=C['green'], fg='white', relief='flat',
                                  cursor='hand2', padx=28, pady=10,
                                  activebackground=C['btn_hover'],
                                  activeforeground='white',
                                  command=self.run)
        self.run_btn.pack(side='left', padx=(0, 12))

        btn_config = tk.Button(btn_row, text='  ⚙️  角色配置  ', font=('Microsoft YaHei', 11),
                               bg=C['orange'], fg='white', relief='flat',
                               cursor='hand2', padx=20, pady=9,
                               activebackground='#d97706',
                               command=self.open_role_editor)
        btn_config.pack(side='left', padx=(0, 12))

        btn_folder = tk.Button(btn_row, text='  📂  打开目录  ', font=('Microsoft YaHei', 11),
                               bg=C['blue'], fg='white', relief='flat',
                               cursor='hand2', padx=20, pady=9,
                               activebackground='#2563eb',
                               command=lambda: os.startfile(SCRIPT_DIR))
        btn_folder.pack(side='left')

        # === 日志区 ===
        c3 = tk.Frame(main, bg=C['card'], highlightthickness=1,
                      highlightbackground=C['border'])
        c3.pack(fill='both', expand=True)

        inner3 = tk.Frame(c3, bg=C['card'])
        inner3.pack(fill='both', expand=True, padx=20, pady=14)

        tk.Label(inner3, text='📝 运行日志',
                 font=('Microsoft YaHei', 13, 'bold'), bg=C['card'],
                 fg=C['text']).pack(anchor='w')

        log_wrap = tk.Frame(inner3, bg=C['log_bg'])
        log_wrap.pack(fill='both', expand=True, pady=(8, 0))

        self.log_txt = tk.Text(log_wrap, font=('Consolas', 10),
                               bg=C['log_bg'], fg=C['log_fg'],
                               insertbackground='white', relief='flat',
                               padx=12, pady=10, wrap='word', state='disabled')
        self.log_txt.pack(side='left', fill='both', expand=True)

        sb = tk.Scrollbar(log_wrap, command=self.log_txt.yview)
        sb.pack(side='right', fill='y')
        self.log_txt.configure(yscrollcommand=sb.set)

        # ---- 底部状态条 ----
        bar = tk.Frame(self.root, bg='#f1f5f9', height=28)
        bar.pack(fill='x', side='bottom')
        bar.pack_propagate(False)
        self.st = tk.Label(bar, text='● 就绪', font=('Microsoft YaHei', 9),
                           bg='#f1f5f9', fg=C['sub'], anchor='w', padx=16)
        self.st.pack(fill='x')

    # ============ 角色标签 ============

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

    def _refresh_file_labels(self):
        def _short(fpath):
            if not fpath: return '（未选择）'
            name = os.path.basename(fpath)
            exists = os.path.exists(fpath)
            icon = '✅' if exists else '❌'
            return f'{icon} {name}'
        self.pf_label.configure(text=_short(self.project_file))
        self.tf_label.configure(text=_short(self.template_file))

    # ============ 文件检查 ============

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
        if not line or 'Traceback' in line or 'Command exited' in line or line.startswith('PS '):
            return
        skip = ['所在位置', 'CategoryInfo', 'FullyQualifiedErrorId', 'NativeCommandError',
                'RemoteException', '~~~~~~~~~~', '按任意键退出', '鎸変换鎰忛', 'Exception', 'EOFError']
        if any(s in line for s in skip): return
        self._log(line)

    # ============ 生成 ============

    def run(self):
        self.run_btn.configure(state='disabled', bg=C['gray'], text='⏳ 生成中...')
        self.st.configure(text='⏳ 正在计算绩效和生成表格...')
        self._log('🔍 开始执行生成流程...')
        self._log('—' * 50)

        def worker():
            try:
                p = subprocess.Popen([PYTHON_EXE, CLI_SCRIPT,
                                      self.project_file, self.template_file],
                                     cwd=SCRIPT_DIR,
                                     stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                     stderr=subprocess.STDOUT, text=True,
                                     encoding='utf-8', errors='replace')
                try: p.stdin.write('\n'); p.stdin.flush()
                except: pass
                for ln in iter(p.readline, ''):
                    self.root.after(0, lambda l=ln: self._filter_line(l))
                p.wait()
            except Exception as e:
                self.root.after(0, lambda: self._log(f'❌ 错误: {e}'))
            self.root.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.run_btn.configure(state='normal', bg=C['green'], text='▶  一键生成提成表')
        self.st.configure(text='● 就绪 · 生成完成')
        self._log('—' * 50)
        self._log('🎉 全部完成！')
        messagebox.showinfo('完成', '✅ 提成表、统计简报和仪表盘已生成完毕！\n\n文件已自动打开。')


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == '__main__':
    main()
