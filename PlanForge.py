import pandas as pd
import numpy as np
import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox, PanedWindow, VERTICAL, Listbox, END
import datetime
import math
import re
import logging
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

class ConfigManager:
    def __init__(self, config_path='config.xlsx'):
        self.config_path = config_path; self.config = {}
        if not os.path.exists(self.config_path): self.create_default_config()
        self.load_config()

    def create_default_config(self):
        default_config = {'PALLET_SIZE': 60, 'LEAD_TIME_DAYS': 2, 'PALLETS_PER_TRUCK': 36, 'MAX_TRUCKS_PER_DAY': 2, 'FONT_SIZE': 11}
        self.save_config(default_config)
        messagebox.showinfo("설정 파일 생성", f"`{self.config_path}` 파일이 생성되었습니다.\n초기 설정값으로 저장되었습니다.")

    def load_config(self):
        try:
            settings_df = pd.read_excel(self.config_path, sheet_name='Settings').set_index('Setting')['Value']
            self.config['PALLET_SIZE'] = int(settings_df.get('PALLET_SIZE', 60))
            self.config['LEAD_TIME_DAYS'] = int(settings_df.get('LEAD_TIME_DAYS', 2))
            self.config['PALLETS_PER_TRUCK'] = int(settings_df.get('PALLETS_PER_TRUCK', 36))
            self.config['MAX_TRUCKS_PER_DAY'] = int(settings_df.get('MAX_TRUCKS_PER_DAY', 2))
            self.config['FONT_SIZE'] = int(settings_df.get('FONT_SIZE', 11))
        except FileNotFoundError: self.create_default_config()
        except Exception as e:
            logging.error(f"Config load error: {e}")
            raise ValueError(f"`{self.config_path}` 로드 중 오류 발생: {e}. 시트 이름이 'Settings'인지 확인하세요.")

    def save_config(self, config_data):
        try:
            with pd.ExcelWriter(self.config_path, engine='openpyxl') as writer:
                settings_df = pd.DataFrame([
                    {'Setting': 'PALLET_SIZE', 'Value': config_data.get('PALLET_SIZE'), 'Description': '하나의 팔레트에 들어가는 제품 수량'},
                    {'Setting': 'LEAD_TIME_DAYS', 'Value': config_data.get('LEAD_TIME_DAYS'), 'Description': '자재 발주 후 도착까지 걸리는 일수 (영업일 기준)'},
                    {'Setting': 'PALLETS_PER_TRUCK', 'Value': config_data.get('PALLETS_PER_TRUCK'), 'Description': '트럭 한 대에 실을 수 있는 최대 팔레트 수'},
                    {'Setting': 'MAX_TRUCKS_PER_DAY', 'Value': config_data.get('MAX_TRUCKS_PER_DAY'), 'Description': '하루에 운행 가능한 최대 트럭 수'},
                    {'Setting': 'FONT_SIZE', 'Value': config_data.get('FONT_SIZE'), 'Description': 'UI 기본 폰트 크기'}
                ])
                settings_df.to_excel(writer, sheet_name='Settings', index=False)
            self.config = config_data
        except Exception as e:
            logging.error(f"Config save error: {e}")
            raise IOError(f"설정 파일 저장 실패: {e}")

class PlanProcessor:
    def __init__(self, config):
        self.config = config
        self.aggregated_plan_df = None
        self.inventory_df = None
        self.simulated_plan_df = None
        self.current_filepath = ""
        self.date_cols = []
        self.inventory_date = None
        self.adjustments = []
        self.item_master_df = None
        self.allowed_models = []
        self.highlight_models = []
        self._load_item_master()

    def _load_item_master(self):
        try:
            item_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'assets', 'Item.csv')
            if not os.path.exists(item_path): raise FileNotFoundError("assets/Item.csv 파일을 찾을 수 없습니다.")
            self.item_master_df = pd.read_csv(item_path)
            if 'Priority' not in self.item_master_df.columns:
                self.item_master_df['Priority'] = range(1, len(self.item_master_df) + 1)
            self.item_master_df.sort_values(by='Priority', inplace=True)
            self.allowed_models = self.item_master_df['Item Code'].tolist()
            self.highlight_models = self.item_master_df[self.item_master_df['Spec'].str.contains('HMC', na=False)]['Item Code'].tolist()
        except Exception as e:
            messagebox.showerror("품목 정보 로드 실패", f"Item.csv 파일 처리 중 오류가 발생했습니다: {e}")
            raise

    def process_plan_file(self):
        try:
            df_raw = pd.read_excel(self.current_filepath, sheet_name='《HCO&DIS》', header=None)
            part_col_index = 11; header_row_index = -1
            for i, row in df_raw.iterrows():
                if len(row) > part_col_index and isinstance(row.iloc[part_col_index], str) and 'cover glass assy' in row.iloc[part_col_index].lower():
                    header_row_index = i; break
            if header_row_index == -1: raise ValueError("헤더 'Cover glass Assy'를 찾을 수 없습니다.")
            df = df_raw.iloc[header_row_index:].copy()
            df.columns = df.iloc[0]; df = df.iloc[1:].rename(columns={df.columns[part_col_index]: 'Model'})
            self.date_cols = sorted([col for col in df.columns if isinstance(col, (datetime.datetime, pd.Timestamp))])
            if not self.date_cols: raise ValueError("파일에서 유효한 날짜 컬럼을 찾을 수 없습니다.")
            df_filtered = df[df['Model'].isin(self.allowed_models)].copy()
            df_filtered.loc[:, self.date_cols] = df_filtered.loc[:, self.date_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
            agg_df = df_filtered.groupby('Model')[self.date_cols].sum()
            reindexed_df = agg_df.reindex(self.allowed_models).fillna(0)
            self.aggregated_plan_df = reindexed_df[reindexed_df.sum(axis=1) > 0].copy()
            return True
        except Exception as e:
            logging.error(f"Plan file processing error: {e}")
            raise

    def load_inventory_from_text(self, text_data):
        data = []; lines = [line.strip() for line in text_data.strip().split('\n') if line.strip()]
        inventory_date = None
        for line in lines:
            date_match = re.search(r'(\d{1,2})/(\d{1,2})', line)
            if date_match:
                month = int(date_match.group(1)); day = int(date_match.group(2)); year = datetime.date.today().year
                inventory_date = datetime.date(year, month, day)
            matches = re.findall(r'(AAA\d+).*?(\d{1,3}(?:,\d{3})*)', line)
            for match in matches:
                model, inventory_str = match
                inventory = int(inventory_str.replace(',', ''))
                data.append({'Model': model, 'Inventory': inventory})
        if not data: raise ValueError("유효한 재고 데이터를 찾을 수 없습니다.")
        self.inventory_df = pd.DataFrame(data).set_index('Model').infer_objects(copy=False)
        self.inventory_date = inventory_date if inventory_date else datetime.date.today()

    def run_simulation(self, adjustments=None):
        self.adjustments = adjustments if adjustments else []
        if self.aggregated_plan_df is None: return
        plan_cols = [col for col in self.aggregated_plan_df.columns if col != 'Status']
        df = self.aggregated_plan_df[plan_cols].copy()
        if self.inventory_df is not None: df = df.join(self.inventory_df, how='left').fillna({'Inventory': 0})
        else: df = df.assign(Inventory=0)
        df['Inventory'] = df['Inventory'].astype(int)
        lead_time = self.config.get('LEAD_TIME_DAYS', 2)
        truck_capacity = self.config.get('PALLETS_PER_TRUCK', 36) * self.config.get('PALLET_SIZE', 60)
        max_trucks = self.config.get('MAX_TRUCKS_PER_DAY', 2); simulated_df = df.copy()
        if not self.date_cols: return
        plan_start = self.date_cols[0].date()
        if self.inventory_date and self.inventory_date > plan_start:
            original_date_count = len(self.date_cols)
            self.date_cols = [d for d in self.date_cols if d.date() >= self.inventory_date]
            if not self.date_cols and original_date_count > 0:
                raise ValueError(f"재고 기준일({self.inventory_date.strftime('%Y-%m-%d')})이 생산 계획의 모든 날짜보다 미래입니다.")
        adjustments_by_date = {adj['date']: [] for adj in self.adjustments}
        for adj in self.adjustments:
            if adj['date'] not in adjustments_by_date: adjustments_by_date[adj['date']] = []
            adjustments_by_date[adj['date']].append(adj)
        shortages = []
        for model in df.index:
            for i, date in enumerate(self.date_cols):
                if (i + lead_time) < len(self.date_cols):
                    future_date = self.date_cols[i + lead_time]
                    production_needed = df.loc[model, future_date]
                    if production_needed > 0:
                        shortages.append({'model': model, 'date_idx': i, 'date': date, 'needed': production_needed, 'urgency': i})
        shortages = sorted(shortages, key=lambda x: (x['urgency'], -x['needed']))
        daily_carryover = {model: 0 for model in df.index}
        for date_idx, date in enumerate(self.date_cols):
            daily_shipments = {model: {f'출고_{t}차_{date.strftime("%m%d")}': 0 for t in range(1, max_trucks + 1)} for model in df.index}
            remaining_capacity_per_truck = [truck_capacity] * max_trucks
            required_dict = {}
            shortages_for_day = [s for s in shortages if s['date_idx'] == date_idx]
            for shortage in shortages_for_day:
                model = shortage['model']
                production_needed = shortage['needed']
                on_hand = simulated_df.loc[model, 'Inventory'] if date_idx == 0 else simulated_df.loc[model, f'재고_{self.date_cols[date_idx-1].strftime("%m%d")}']
                future_date_idx = date_idx + lead_time
                intermediate_cols = self.date_cols[date_idx:future_date_idx]
                sum_intermediate_prod = df.loc[model, intermediate_cols].sum()
                total_needed = production_needed + daily_carryover[model] + sum_intermediate_prod
                required = max(0, total_needed - on_hand)
                simulated_df.loc[model, f'수요_{date.strftime("%m%d")}'] = production_needed
                simulated_df.loc[model, f'이월_{date.strftime("%m%d")}'] = daily_carryover[model]
                simulated_df.loc[model, f'요구_{date.strftime("%m%d")}'] = required
                if required > 0: required_dict[model] = required
            sorted_models = sorted(required_dict, key=required_dict.get, reverse=True)
            for model in sorted_models:
                remaining_to_ship = required_dict[model]
                for truck_num in range(max_trucks):
                    if remaining_capacity_per_truck[truck_num] > 0:
                        shipment = min(remaining_to_ship, remaining_capacity_per_truck[truck_num])
                        daily_shipments[model][f'출고_{truck_num+1}차_{date.strftime("%m%d")}'] += shipment
                        remaining_capacity_per_truck[truck_num] -= shipment
                        remaining_to_ship -= shipment
                        if remaining_to_ship == 0: break
                daily_carryover[model] = remaining_to_ship
            if any(cap > 0 for cap in remaining_capacity_per_truck):
                future_shortages = [s for s in shortages if s['date_idx'] > date_idx]
                future_shortages.sort(key=lambda x: (x['urgency'], -x['needed']))
                for shortage in future_shortages:
                    model = shortage['model']
                    future_needed = shortage['needed']
                    if future_needed > 0:
                        remaining_proactive = future_needed
                        for truck_num in range(max_trucks):
                            if remaining_capacity_per_truck[truck_num] > 0:
                                proactive = min(remaining_proactive, remaining_capacity_per_truck[truck_num])
                                daily_shipments[model][f'출고_{truck_num+1}차_{date.strftime("%m%d")}'] += proactive
                                remaining_capacity_per_truck[truck_num] -= proactive
                                remaining_proactive -= proactive
                                if remaining_proactive <= 0: break
                    if all(cap == 0 for cap in remaining_capacity_per_truck): break
            for model in df.index:
                today_production = df.loc[model, date]
                inventory_adjustment = 0
                for adj in adjustments_by_date.get(date.date(), []):
                    if adj['model'] == model:
                        if adj['type'] == '수요': today_production += adj['qty']
                        elif adj['type'] == '재고': inventory_adjustment += adj['qty']
                total_shipped = sum(daily_shipments[model].values())
                on_hand = simulated_df.loc[model, 'Inventory'] if date_idx == 0 else simulated_df.loc[model, f'재고_{self.date_cols[date_idx-1].strftime("%m%d")}']
                new_inventory = max(0, on_hand - today_production + total_shipped + inventory_adjustment)
                simulated_df.loc[model, f'재고_{date.strftime("%m%d")}'] = new_inventory
                for k, v in daily_shipments[model].items():
                    simulated_df.loc[model, k] = v
        self.simulated_plan_df = simulated_df.astype(float).fillna(0).astype(int).infer_objects(copy=False)

class AdjustmentDialog(ctk.CTkToplevel):
    def __init__(self, parent, models):
        super().__init__(parent)
        self.models = models
        self.adjustments = []
        self.result = None
        self.title("수동 조정 입력"); self.geometry("600x450")
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ctk.CTkLabel(input_frame, text="모델:").grid(row=0, column=0, padx=5, pady=5)
        self.model_combo = ctk.CTkComboBox(input_frame, values=self.models, width=150)
        self.model_combo.grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkLabel(input_frame, text="날짜 (YYYY-MM-DD):").grid(row=0, column=2, padx=5, pady=5)
        self.date_entry = ctk.CTkEntry(input_frame, placeholder_text=datetime.date.today().strftime('%Y-%m-%d'))
        self.date_entry.grid(row=0, column=3, padx=5, pady=5)
        ctk.CTkLabel(input_frame, text="수량:").grid(row=1, column=0, padx=5, pady=5)
        self.qty_entry = ctk.CTkEntry(input_frame)
        self.qty_entry.grid(row=1, column=1, padx=5, pady=5)
        ctk.CTkLabel(input_frame, text="타입:").grid(row=1, column=2, padx=5, pady=5)
        self.type_combo = ctk.CTkComboBox(input_frame, values=['재고', '수요'])
        self.type_combo.grid(row=1, column=3, padx=5, pady=5)
        ctk.CTkButton(input_frame, text="추가", command=self.add_adjustment).grid(row=1, column=4, padx=10, pady=5)
        self.listbox = Listbox(self, height=10)
        self.listbox.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=10, pady=10, sticky="e")
        ctk.CTkButton(button_frame, text="확인", command=self.ok_event).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="취소", command=self.cancel_event, fg_color="gray").pack(side="left")
        self.transient(parent); self.grab_set()

    def add_adjustment(self):
        model = self.model_combo.get()
        date_str = self.date_entry.get() or self.date_entry.cget("placeholder_text")
        qty_str = self.qty_entry.get()
        adj_type = self.type_combo.get()
        if not all([model, date_str, qty_str, adj_type]):
            messagebox.showwarning("입력 오류", "모든 필드를 채워주세요.", parent=self)
            return
        try:
            adj_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
            quantity = int(qty_str)
        except ValueError:
            messagebox.showwarning("형식 오류", "날짜는 'YYYY-MM-DD', 수량은 숫자로 입력해야 합니다.", parent=self)
            return
        adj = {'model': model, 'date': adj_date, 'qty': quantity, 'type': adj_type}
        self.adjustments.append(adj)
        self.listbox.insert(END, f"{adj['date']}, {adj['model']}, {adj['qty']}, {adj['type']}")
        self.qty_entry.delete(0, END)

    def ok_event(self):
        self.result = self.adjustments
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()

# 추가된 MultilineInputDialog 클래스
class MultilineInputDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, prompt):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x300")
        self.result = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text=prompt).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.textbox = ctk.CTkTextbox(self, width=380, height=150)
        self.textbox.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="e")

        ctk.CTkButton(button_frame, text="확인", command=self.ok_event).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="취소", command=self.cancel_event, fg_color="gray").pack(side="left", padx=5)

        self.transient(parent)
        self.grab_set()

    def ok_event(self):
        self.result = self.textbox.get("1.0", "end-1c")
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()

class ProductionPlannerApp(ctk.CTk):
    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.processor = PlanProcessor(self.config_manager.config)
        self.current_step = 0; self.current_file = "파일이 로드되지 않았습니다.";
        self.base_font_size = self.config_manager.config.get('FONT_SIZE', 11)
        self.font_big_bold = ctk.CTkFont(size=20, weight="bold")
        self.font_normal = ctk.CTkFont(size=self.base_font_size); self.font_small = ctk.CTkFont(size=self.base_font_size - 1)
        self.font_bold = ctk.CTkFont(size=self.base_font_size, weight="bold"); self.font_italic = ctk.CTkFont(size=self.base_font_size, slant="italic")
        self.font_kpi = ctk.CTkFont(size=14, weight="bold"); self.font_header = ctk.CTkFont(size=self.base_font_size + 1, weight="bold")
        self.title("PlanForge Pro - 출고계획 시스템 v22.0 (최종 개선)"); self.geometry("1800x1000")
        ctk.set_appearance_mode("Light"); ctk.set_default_color_theme("blue")
        self.grid_rowconfigure(0, weight=1); self.grid_columnconfigure(1, weight=1)
        self.create_widgets(); self.update_status_bar(); self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind_all("<Control-MouseWheel>", self.on_mouse_wheel_zoom)
        self.inventory_text_backup = None
        self.after_ids = []

    def on_closing(self):
        try:
            for after_id in self.after_ids: self.after_cancel(after_id)
            self.after_ids = []
            self.unbind_all("<Control-MouseWheel>")
            plt.close('all')
            if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
                self.destroy()
        except Exception as e: 
            logging.error(f"Closing error: {e}")
            self.destroy()

    def on_mouse_wheel_zoom(self, event): self.set_font_size(self.base_font_size + (1 if event.delta > 0 else -1))

    def create_widgets(self):
        sidebar_frame = ctk.CTkFrame(self, width=280, corner_radius=0)
        sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew"); sidebar_frame.grid_rowconfigure(6, weight=1)
        self.sidebar_title = ctk.CTkLabel(sidebar_frame, text="PlanForge Pro", font=self.font_big_bold)
        self.sidebar_title.pack(pady=20)
        self.step1_button = ctk.CTkButton(sidebar_frame, text="1. 생산계획 불러오기", command=self.run_step1_aggregate, font=self.font_normal)
        self.step1_button.pack(fill='x', padx=20, pady=5)
        self.step2_button = ctk.CTkButton(sidebar_frame, text="2. 재고 반영 및 계획 시뮬레이션", command=self.run_step2_simulation, state="disabled", font=self.font_normal)
        self.step2_button.pack(fill='x', padx=20, pady=5)
        self.step3_button = ctk.CTkButton(sidebar_frame, text="3. 수동 조정 적용", command=self.run_step3_adjustments, state="disabled", font=self.font_normal)
        self.step3_button.pack(fill='x', padx=20, pady=5)
        self.step4_button = ctk.CTkButton(sidebar_frame, text="4. 계획 내보내기 (Excel)", command=self.export_to_excel, state="disabled", font=self.font_normal)
        self.step4_button.pack(fill='x', padx=20, pady=5)
        font_frame = ctk.CTkFrame(sidebar_frame, fg_color="transparent")
        font_frame.pack(fill='x', padx=20, pady=10)
        self.font_size_title_label = ctk.CTkLabel(font_frame, text="폰트 크기:", font=self.font_normal)
        self.font_size_title_label.pack(side="left")
        self.font_minus_button = ctk.CTkButton(font_frame, text="-", width=30, command=lambda: self.change_font_size(-1), font=self.font_normal)
        self.font_minus_button.pack(side="left", padx=5)
        self.font_size_label = ctk.CTkLabel(font_frame, text=str(self.base_font_size), font=self.font_normal, cursor="hand2")
        self.font_size_label.pack(side="left")
        self.font_size_label.bind("<Button-1>", self.prompt_for_font_size)
        self.font_plus_button = ctk.CTkButton(font_frame, text="+", width=30, command=lambda: self.change_font_size(1), font=self.font_normal)
        self.font_plus_button.pack(side="left", padx=5)
        settings_frame = ctk.CTkFrame(sidebar_frame, fg_color="transparent")
        settings_frame.pack(fill='x', expand=True, padx=20, pady=20)
        self.settings_title_label = ctk.CTkLabel(settings_frame, text="시스템 설정", font=self.font_bold)
        self.settings_title_label.pack()
        self.settings_entries = {}
        settings_map = {'팔레트당 수량': 'PALLET_SIZE', '리드타임 (일)': 'LEAD_TIME_DAYS', '트럭당 팔레트 수': 'PALLETS_PER_TRUCK', '일일 최대 차수': 'MAX_TRUCKS_PER_DAY'}
        self.setting_labels = []
        for label_text, key in settings_map.items():
            frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
            frame.pack(fill='x', pady=2)
            label = ctk.CTkLabel(frame, text=label_text, width=120, font=self.font_normal)
            label.pack(side='left'); self.setting_labels.append(label)
            entry = ctk.CTkEntry(frame, font=self.font_normal); entry.pack(side='left', fill='x', expand=True)
            self.settings_entries[key] = entry
        self.load_settings_to_gui()
        self.save_settings_button = ctk.CTkButton(sidebar_frame, text="설정 저장 및 재계산", command=self.save_settings_and_recalculate, fg_color="#1F6AA5", font=self.font_normal)
        self.save_settings_button.pack(fill='x', padx=20, pady=10, side='bottom')
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        main_frame.grid_rowconfigure(2, weight=1); main_frame.grid_columnconfigure(0, weight=1)
        search_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky='ew', pady=(0,5))
        self.search_label = ctk.CTkLabel(search_frame, text="품목 검색:", font=self.font_normal)
        self.search_label.pack(side='left', padx=(0,5))
        self.search_entry = ctk.CTkEntry(search_frame, font=self.font_normal)
        self.search_entry.pack(side='left', fill='x', expand=True)
        self.search_entry.bind("<KeyRelease>", self.filter_grid)
        self.kpi_frame = ctk.CTkFrame(main_frame, fg_color="#EAECEE", corner_radius=5)
        self.kpi_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.kpi_frame.grid_columnconfigure((0,1,2), weight=1)
        self.lbl_models_found = ctk.CTkLabel(self.kpi_frame, text="처리된 모델 수: -", font=self.font_kpi)
        self.lbl_models_found.grid(row=0, column=0, padx=10, pady=10)
        self.lbl_total_quantity = ctk.CTkLabel(self.kpi_frame, text="총생산량: -", font=self.font_kpi)
        self.lbl_total_quantity.grid(row=0, column=1, padx=10, pady=10)
        self.lbl_date_range = ctk.CTkLabel(self.kpi_frame, text="계획 기간: -", font=self.font_kpi)
        self.lbl_date_range.grid(row=0, column=2, padx=10, pady=10)
        paned_window = PanedWindow(main_frame, orient=VERTICAL, sashwidth=8, bg="#EAECEE", showhandle=True, handlepad=20, handlesize=10)
        paned_window.grid(row=2, column=0, sticky="nsew")
        master_container = ctk.CTkFrame(paned_window, fg_color="transparent")
        master_container.grid_rowconfigure(0, weight=1); master_container.grid_columnconfigure(0, weight=1)
        self.master_frame = ctk.CTkScrollableFrame(master_container, label_text="개요: 전체 생산계획", label_font=self.font_bold)
        self.master_frame.grid(row=0, column=0, sticky="nsew")
        paned_window.add(master_container, height=600)
        detail_container = ctk.CTkFrame(paned_window, fg_color="transparent")
        detail_container.grid_rowconfigure(0, weight=1); detail_container.grid_columnconfigure(0, weight=1)
        self.detail_frame = ctk.CTkScrollableFrame(detail_container, label_text="상세: 선택된 모델의 출고 시뮬레이션", label_font=self.font_bold)
        self.detail_frame.grid(row=0, column=0, sticky="nsew")
        paned_window.add(detail_container)
        self.status_bar = ctk.CTkLabel(self, text="", anchor="w", font=self.font_normal)
        self.status_bar.grid(row=1, column=1, sticky="ew", padx=10, pady=(0, 5))

    def prompt_for_font_size(self, event=None):
        dialog = ctk.CTkInputDialog(text="새로운 폰트 크기를 입력하세요 (8-30):", title="폰트 크기 변경")
        new_size_str = dialog.get_input()
        if new_size_str:
            try:
                new_size = int(new_size_str)
                if not (8 <= new_size <= 30): messagebox.showwarning("입력 오류", "폰트 크기는 8과 30 사이의 숫자여야 합니다.")
                else: self.set_font_size(new_size)
            except (ValueError, TypeError): messagebox.showerror("입력 오류", "유효한 숫자를 입력해주세요.")

    def set_font_size(self, new_size):
        if new_size == self.base_font_size: return
        self.base_font_size = new_size
        self.font_normal.configure(size=self.base_font_size); self.font_small.configure(size=self.base_font_size - 1)
        self.font_bold.configure(size=self.base_font_size, weight="bold"); self.font_italic.configure(size=self.base_font_size, slant="italic")
        self.font_header.configure(size=self.base_font_size + 1, weight="bold")
        self.update_static_fonts()

    def change_font_size(self, delta):
        new_size = max(8, min(30, self.base_font_size + delta))
        self.set_font_size(new_size)

    def update_static_fonts(self):
        self.sidebar_title.configure(font=self.font_big_bold)
        self.step1_button.configure(font=self.font_normal); self.step2_button.configure(font=self.font_normal)
        self.step3_button.configure(font=self.font_normal); self.step4_button.configure(font=self.font_normal)
        self.font_size_title_label.configure(font=self.font_normal); self.font_minus_button.configure(font=self.font_normal)
        self.font_size_label.configure(font=self.font_normal, text=str(self.base_font_size))
        self.font_plus_button.configure(font=self.font_normal)
        self.settings_title_label.configure(font=self.font_bold)
        for label in self.setting_labels: label.configure(font=self.font_normal)
        for entry in self.settings_entries.values(): entry.configure(font=self.font_normal)
        self.save_settings_button.configure(font=self.font_normal)
        self.search_label.configure(font=self.font_normal)
        self.search_entry.configure(font=self.font_normal)
        self.lbl_models_found.configure(font=self.font_kpi); self.lbl_total_quantity.configure(font=self.font_kpi)
        self.lbl_date_range.configure(font=self.font_kpi)
        self.master_frame.configure(label_font=self.font_bold); self.detail_frame.configure(label_font=self.font_bold)
        self.filter_grid()
        if self.current_step >= 2 and hasattr(self, 'last_selected_model'):
            self.populate_detail_view(self.last_selected_model)

    def populate_master_grid(self, df_to_show):
        for widget in self.master_frame.winfo_children(): widget.destroy()
        if df_to_show is None: return
        max_trucks = self.config_manager.config.get('MAX_TRUCKS_PER_DAY', 2)
        plan_cols = self.processor.date_cols
        headers = ['Model'] + [d.strftime('%m-%d') for d in plan_cols]
        col_widths = [140] + [70] * len(plan_cols)
        for i, width in enumerate(col_widths): self.master_frame.grid_columnconfigure(i, minsize=width)
        for c, h_text in enumerate(headers):
             ctk.CTkLabel(self.master_frame, text=h_text, font=self.font_header, anchor="center").grid(row=0, column=c, sticky="ew", padx=1)
        df_display = df_to_show.reset_index()
        for r, row_data in df_display.iterrows():
            model = row_data['Model']
            is_highlighted = model in self.processor.highlight_models
            bg_color = "transparent"
            if self.current_step >= 2 : bg_color = "#E6F3E6"
            if is_highlighted: bg_color = "#D6EAF8"
            row_widgets = []
            lbl_model = ctk.CTkLabel(self.master_frame, text=model, fg_color=bg_color, font=self.font_normal, anchor="w", padx=5)
            lbl_model.grid(row=r+1, column=0, sticky="ew"); row_widgets.append(lbl_model)
            for i, date_col in enumerate(plan_cols):
                if self.current_step < 2:
                    val = row_data.get(date_col, 0)
                else:
                    date_str = date_col.strftime("%m%d")
                    val = sum(row_data.get(f'출고_{t}차_{date_str}', 0) for t in range(1, max_trucks + 1))
                text = f"{val:,.0f}" if val else "0"
                lbl_data = ctk.CTkLabel(self.master_frame, text=text, fg_color=bg_color, font=self.font_normal)
                lbl_data.grid(row=r+1, column=i+1, sticky="ew"); row_widgets.append(lbl_data)
            for widget in row_widgets: widget.bind("<Button-1>", lambda e, m=model: self.on_row_clicked(m))
        if not df_to_show.empty:
            ctk.CTkFrame(self.master_frame, height=1, fg_color="lightgray").grid(row=len(df_display)+1, column=0, columnspan=len(headers), sticky='ew', pady=4)
            ctk.CTkLabel(self.master_frame, text="합계", font=self.font_bold, anchor="w", padx=5).grid(row=len(df_display)+2, column=0, sticky="ew")
            if self.current_step < 2:
                totals = df_to_show[plan_cols].sum()
            else:
                totals = pd.Series(index=plan_cols, dtype=int)
                for date_col in plan_cols:
                    date_str = date_col.strftime("%m%d")
                    totals[date_col] = sum(df_to_show[f'출고_{t}차_{date_str}'].sum() for t in range(1, max_trucks + 1))
            for i, date_col in enumerate(plan_cols):
                total_val = totals[date_col]
                ctk.CTkLabel(self.master_frame, text=f"{total_val:,.0f}", font=self.font_bold).grid(row=len(df_display)+2, column=i+1, sticky="ew")

    def populate_detail_view(self, model_name):
        for widget in self.detail_frame.winfo_children(): widget.destroy()
        df = self.processor.simulated_plan_df
        if df is None or model_name not in df.index: 
            ctk.CTkLabel(self.detail_frame, text=f"'{model_name}'에 대한 시뮬레이션 데이터를 찾을 수 없습니다.", font=self.font_normal).pack(pady=20)
            return
        row_data = df.loc[model_name]
        self.detail_frame.configure(label_text=f"상세: '{model_name}' 출고 시뮬레이션")
        date_cols = self.processor.date_cols; max_trucks = self.config_manager.config.get('MAX_TRUCKS_PER_DAY', 2)
        headers = ['항목'] + [d.strftime('%m-%d') for d in date_cols]
        for c, h in enumerate(headers):
            self.detail_frame.grid_columnconfigure(c, minsize=120 if c == 0 else 85)
            ctk.CTkLabel(self.detail_frame, text=h, font=self.font_header, anchor="center").grid(row=0, column=c, padx=2, pady=1, sticky="ew")
        current_row = 1
        def draw_row(name, key_prefix, options={}, is_date_key=False, is_initial=False):
            nonlocal current_row
            font_opts = {k: v for k, v in options.items() if k in ['weight', 'slant']}
            label_opts = {k: v for k, v in options.items() if k not in ['weight', 'slant']}
            label_font = self.font_italic if 'slant' in font_opts else self.font_normal
            if 'weight' in font_opts and font_opts['weight'] == 'bold': label_font = self.font_bold
            data_font = self.font_bold if 'weight' in font_opts else self.font_normal
            ctk.CTkLabel(self.detail_frame, text=name, anchor="w", font=label_font, **label_opts).grid(row=current_row, column=0, sticky="w", padx=5)
            if is_initial:
                val = row_data.get(key_prefix, 0)
                ctk.CTkLabel(self.detail_frame, text=f"{val:,.0f}", font=data_font, **label_opts).grid(row=current_row, column=1, sticky="w")
            else:
                for c, date_col in enumerate(date_cols):
                    val = row_data.get(date_col, 0) if is_date_key else row_data.get(f'{key_prefix}{date_col.strftime("%m%d")}', 0)
                    ctk.CTkLabel(self.detail_frame, text=f"{val:,.0f}", font=data_font, **label_opts, anchor="e").grid(row=current_row, column=c+1, padx=5, sticky="ew")
            current_row += 1
        def draw_separator():
            nonlocal current_row
            ctk.CTkFrame(self.detail_frame, height=1, fg_color="lightgray").grid(row=current_row, column=0, columnspan=len(headers), sticky="ew", pady=4); current_row += 1
        draw_row("초기 재고", "Inventory", options={'weight':'bold'}, is_initial=True)
        draw_row("당일 생산", "", options={}, is_date_key=True); draw_separator()
        draw_row("L/T 적용 수요", "수요_", options={'slant':'italic'})
        draw_row("전일 이월량", "이월_", options={'slant':'italic'})
        draw_row("총 출고 요구량", "요구_", options={'weight':'bold'}); draw_separator()
        for i in range(1, max_trucks + 1):
            draw_row(f"{i}차 출고", f"출고_{i}차_", options={'weight':'bold', 'text_color':"#2E86C1"})
        draw_separator()
        draw_row("일일 재고", "재고_", options={'weight':'bold'})
        fig, ax = plt.subplots(figsize=(8, 4))
        inventory_vals = [row_data.get(f'재고_{d.strftime("%m%d")}', 0) for d in date_cols]
        labels = [d.strftime('%m-%d') for d in date_cols]
        ax.plot(labels, inventory_vals, marker='o')
        ax.set_title('일일 재고 추이'); ax.set_xlabel('날짜'); ax.set_ylabel('재고량')
        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, rotation=45)
        canvas = FigureCanvasTkAgg(fig, master=self.detail_frame)
        canvas.draw(); canvas.get_tk_widget().grid(row=current_row, column=0, columnspan=len(headers), pady=10)

    def run_step1_aggregate(self):
        file_path = filedialog.askopenfilename(title="생산계획 엑셀 파일 선택", filetypes=(("Excel", "*.xlsx *.xls"),))
        if not file_path: return
        try:
            self.processor.current_filepath = file_path; self.processor.process_plan_file()
            self.current_file = os.path.basename(file_path); self.current_step = 1
            if self.processor.aggregated_plan_df.empty:
                messagebox.showinfo("정보", "처리할 생산 계획 데이터가 없습니다.")
                return
            plan_cols = self.processor.date_cols; df = self.processor.aggregated_plan_df
            models_found = len(df.index); total_qty = df[plan_cols].sum().sum()
            date_range = f"{plan_cols[0].strftime('%y/%m/%d')} ~ {plan_cols[-1].strftime('%y/%m/%d')}"
            self.lbl_models_found.configure(text=f"처리된 모델 수: {models_found} 개"); self.lbl_total_quantity.configure(text=f"총생산량: {total_qty:,.0f} 개")
            self.lbl_date_range.configure(text=f"계획 기간: {date_range}")
            self.master_frame.configure(label_text="개요: 전체 생산계획")
            self.filter_grid(); [widget.destroy() for widget in self.detail_frame.winfo_children()]
            self.update_status_bar("1단계: 생산계획 집계 완료"); self.step2_button.configure(state="normal")
        except Exception as e: messagebox.showerror("1단계 파일 처리 실패", f"{e}")

    def check_shipment_capacity(self):
        df = self.processor.simulated_plan_df
        if df is None or not self.processor.date_cols: return
        date_cols = self.processor.date_cols
        max_trucks = self.config_manager.config.get('MAX_TRUCKS_PER_DAY', 2)
        truck_capacity = self.config_manager.config.get('PALLETS_PER_TRUCK', 36) * self.config_manager.config.get('PALLET_SIZE', 60)
        messages = []
        for date in date_cols:
            date_str = date.strftime("%m%d")
            for truck_num in range(1, max_trucks + 1):
                col = f'출고_{truck_num}차_{date_str}'
                if col not in df.columns: continue
                total_shipped = df[col].sum()
                if total_shipped > truck_capacity:
                    messages.append(f"{date.strftime('%m-%d')} {truck_num}차: 출고량 {total_shipped} > 용량 {truck_capacity}.")
        if messages:
            messagebox.showwarning("출고 용량 초과", "\n".join(messages))

    def run_step2_simulation(self):
        dialog = MultilineInputDialog(self, title="재고 데이터 입력", prompt="엑셀에서 복사한 재고 데이터를 아래에 붙여넣으세요:")
        self.wait_window(dialog); pasted_text = dialog.result
        if not pasted_text: return
        try:
            self.inventory_text_backup = pasted_text
            self.processor.load_inventory_from_text(pasted_text)
            self.processor.run_simulation(adjustments=None)
            self.current_step = 2
            total_ship = self.processor.simulated_plan_df[[col for col in self.processor.simulated_plan_df.columns if isinstance(col, str) and col.startswith('출고_')]].sum().sum()
            self.lbl_total_quantity.configure(text=f"총출고량: {total_ship:,.0f} 개")
            self.master_frame.configure(label_text="개요: 전체 출고계획")
            self.filter_grid(); [widget.destroy() for widget in self.detail_frame.winfo_children()]
            self.update_status_bar("2단계: 출고 계획 시뮬레이션 완료.")
            self.step3_button.configure(state="normal")
            self.step4_button.configure(state="normal")
            self.check_shipment_capacity()
        except Exception as e: messagebox.showerror("2단계 시뮬레이션 실패", f"{e}")

    def run_step3_adjustments(self):
        if self.current_step < 2: 
            messagebox.showwarning("오류", "먼저 2단계(재고 반영)를 실행해야 합니다.")
            return
        dialog = AdjustmentDialog(self, models=self.processor.allowed_models)
        self.wait_window(dialog)
        adjustments = dialog.result
        if adjustments is None: return
        try:
            self.processor.load_inventory_from_text(self.inventory_text_backup)
            self.processor.run_simulation(adjustments=adjustments)
            self.current_step = 3
            total_ship = self.processor.simulated_plan_df[[col for col in self.processor.simulated_plan_df.columns if isinstance(col, str) and col.startswith('출고_')]].sum().sum()
            self.lbl_total_quantity.configure(text=f"총출고량: {total_ship:,.0f} 개")
            self.master_frame.configure(label_text="개요: 전체 출고계획 (조정됨)")
            self.filter_grid()
            [widget.destroy() for widget in self.detail_frame.winfo_children()]
            self.update_status_bar("3단계: 수동 조정 적용 완료.")
            self.check_shipment_capacity()
        except Exception as e: 
            messagebox.showerror("3단계 조정 실패", f"{e}")

    def export_to_excel(self):
        if self.processor.simulated_plan_df is None: return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel", "*.xlsx"),))
        if not file_path: return
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                self.processor.simulated_plan_df.to_excel(writer, sheet_name='Full Plan')
                max_trucks = self.config_manager.config.get('MAX_TRUCKS_PER_DAY', 2)
                date_cols = self.processor.date_cols
                models = self.processor.simulated_plan_df.index
                for truck_num in range(1, max_trucks + 1):
                    sheet_name = f'{truck_num}차 출고'
                    df_round = pd.DataFrame(index=models, columns=date_cols)
                    for date in date_cols:
                        date_str = date.strftime("%m%d")
                        col = f'출고_{truck_num}차_{date_str}'
                        if col in self.processor.simulated_plan_df.columns:
                            df_round[date] = self.processor.simulated_plan_df[col]
                    df_round.columns = [d.strftime('%m-%d') for d in date_cols]
                    df_round.to_excel(writer, sheet_name=sheet_name)
            messagebox.showinfo("내보내기 성공", f"계획이 {file_path}로 저장되었습니다.")
        except Exception as e: 
            logging.error(f"Export error: {e}")
            messagebox.showerror("내보내기 실패", f"{e}")

    def on_row_clicked(self, model_name):
        self.last_selected_model = model_name
        if self.current_step < 2: self.update_status_bar(f"'{model_name}'의 상세 정보를 보려면 먼저 2단계 시뮬레이션을 실행하세요."); return
        self.populate_detail_view(model_name)

    def filter_grid(self, event=None):
        if self.current_step < 2:
            df_to_show = self.processor.aggregated_plan_df
        else:
            df_to_show = self.processor.simulated_plan_df
        if df_to_show is None: return
        search_term = self.search_entry.get().lower()
        if search_term: 
            df_to_show = df_to_show[df_to_show.index.str.lower().str.contains(search_term)]
        self.populate_master_grid(df_to_show)

    def update_status_bar(self, message="준비 완료"): self.status_bar.configure(text=f"현재 파일: {self.current_file} | 상태: {message}")

    def load_settings_to_gui(self):
        for key, entry_widget in self.settings_entries.items():
            entry_widget.delete(0, 'end'); entry_widget.insert(0, str(self.config_manager.config.get(key, '')))

    def save_settings_and_recalculate(self):
        new_config = self.config_manager.config.copy()
        try:
            for key, entry_widget in self.settings_entries.items(): new_config[key] = int(entry_widget.get())
            self.config_manager.save_config(new_config); self.processor.config = new_config
            if self.current_step >= 1: self.processor.process_plan_file()
            if self.current_step >= 2:
                if self.inventory_text_backup:
                    self.processor.load_inventory_from_text(self.inventory_text_backup)
                self.processor.run_simulation(adjustments=self.processor.adjustments)
                total_ship = self.processor.simulated_plan_df[[col for col in self.processor.simulated_plan_df.columns if isinstance(col, str) and col.startswith('출고_')]].sum().sum()
                self.lbl_total_quantity.configure(text=f"총출고량: {total_ship:,.0f} 개")
                self.master_frame.configure(label_text="개요: 전체 출고계획 (설정 변경됨)")
                self.check_shipment_capacity()
            self.filter_grid()
            messagebox.showinfo("성공", "설정이 저장되었고 현재 단계까지 재계산되었습니다.")
        except Exception as e: 
            logging.error(f"Settings save and recalc error: {e}")
            messagebox.showerror("오류", f"설정 저장 및 재계산 실패: {e}")

if __name__ == "__main__":
    try:
        config_manager = ConfigManager()
        app = ProductionPlannerApp(config_manager)
        app.mainloop()
    except Exception as e: 
        logging.critical(f"Fatal error: {e}")
        messagebox.showerror("치명적 오류", f"프로그램 실행에 실패했습니다.\n{e}")