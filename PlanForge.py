import pandas as pd
import numpy as np
import os
import sys
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, PanedWindow, VERTICAL, HORIZONTAL, Listbox, END, Menu, simpledialog
import datetime
import math
import re
import logging
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkcalendar import Calendar, DateEntry
from tkinter import Toplevel, ttk
import requests
import zipfile
import subprocess
import threading
from queue import Queue, Empty

# --- Business Logic & Workflow ---
# 1. 목표 (Goal):
#    - 고객사의 생산 계획과 현재고를 바탕으로, 재고 부족(Stock-out) 및 과잉을 최소화하는 일일 최적 부품 납품 수량을 계산한다.
#
# 2. 핵심 프로세스 (Core Process):
#    - 입력 (Input): 고객사 주간 생산 계획(Excel), 고객사 창고 부품 재고(Text 또는 Excel/CSV 파일)
#    - 출력 (Output): 일자별, 모델별, 트럭 차수별 납품 계획
#
# 3. 주요 제약 조건 (Key Constraints):
#    - 납품 시점 (Delivery Timing): 고객이 특정일(D-day)에 생산할 부품은, '적어도' 그 전날(D-1)까지는 고객사 창고에 도착해야 한다.
#    - 출고 단위 (Shipment Unit): 1 트럭 = 36 팔레트, 1 팔레트 = 60 개. 따라서 1 트럭의 최대 적재량은 2,160개이다.
#    - 출고 빈도 (Shipment Frequency): 하루 최대 2회 출고를 기본으로 하나, 필요시 3차, 4차 출고도 고려할 수 있다 (설정 가능).
#
# 4. 출고 결정 로직 (Shipment Decision Logic):
#    - 우선순위 기반 적재 (Priority-Based Loading): 'Item.csv'에 정의된 우선순위에 따라 긴급한 품목부터 트럭에 적재한다.
#
#    - 이중 예측 기반 필요량 산출 (Dual-Horizon Requirement Calculation):
#        - 단기 예측 (Short-Term): '리드타임'을 기반으로 당장 긴급하게 필요한 물량을 계산한다 (예: 2-3일).
#        - 장기 예측 (Long-Term): 더 긴 미래(예: 7일)의 총생산량을 함께 예측하여, 갑작스러운 생산량 폭증에 대비한 선제적 납품 물량을 계산한다.
#        - 최종 결정: 단기/장기 예측 중 더 많은 수량을 요구하는 쪽을 기준으로 당일 필요량을 결정하여, 미래의 결품 위기를 사전에 방지한다.
#
#    - 계획 건전성 검사 (Plan Health Check): 필수 출고량을 당일 운송 용량 내에서 해결할 수 없는 경우, 이를 '계획 실패'로 간주하고 사용자에게 명확히 경고한다.
# -----------------------------------------

# ===================================================================
# GitHub 자동 업데이트 설정
# ===================================================================
REPO_OWNER = "Your-GitHub-Username"
REPO_NAME = "PlanForge-Repository-Name"
CURRENT_VERSION = "v1.0.0" # 버전 업데이트
# ===================================================================

# 자동 업데이트 기능 함수
def check_for_updates(repo_owner: str, repo_name: str, current_version: str):
    logging.info("Checking for updates...")
    try:
        api_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
        response = requests.get(api_url, timeout=5)
        response.raise_for_status()
        latest_release_data = response.json()
        latest_version = latest_release_data['tag_name']
        logging.info(f"Current version: {current_version}, Latest version: {latest_version}")
        clean_current = current_version.lower().lstrip('v').split('-')[0]
        clean_latest = latest_version.lower().lstrip('v').split('-')[0]
        if clean_latest > clean_current:
            for asset in latest_release_data['assets']:
                if asset['name'].endswith('.zip'):
                    return asset['browser_download_url'], latest_version
    except requests.exceptions.RequestException as e:
        logging.error(f"Update check failed: {e}")
    return None, None

def download_and_apply_update(url: str):
    try:
        logging.info(f"Downloading update from: {url}")
        temp_dir = os.environ.get("TEMP", "C:\\Temp")
        zip_path = os.path.join(temp_dir, "update.zip")
        response = requests.get(url, stream=True, timeout=120)
        response.raise_for_status()
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192): f.write(chunk)
        logging.info("Download complete.")
        temp_update_folder = os.path.join(temp_dir, "temp_update")
        if os.path.exists(temp_update_folder):
            import shutil
            shutil.rmtree(temp_update_folder)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_update_folder)
        os.remove(zip_path)
        logging.info(f"Extracted update to: {temp_update_folder}")
        application_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        updater_script_path = os.path.join(application_path, "updater.bat")
        extracted_content = os.listdir(temp_update_folder)
        new_program_folder_path = temp_update_folder
        if len(extracted_content) == 1 and os.path.isdir(os.path.join(temp_update_folder, extracted_content[0])):
            new_program_folder_path = os.path.join(temp_update_folder, extracted_content[0])
        with open(updater_script_path, "w", encoding='utf-8') as bat_file:
            bat_file.write(fr"""@echo off
chcp 65001 > nul
echo.
echo ==========================================================
echo  프로그램을 업데이트합니다. 이 창을 닫지 마세요.
echo ==========================================================
echo.
echo 잠시 후 프로그램이 자동으로 종료됩니다...
timeout /t 3 /nobreak > nul
taskkill /F /IM "{os.path.basename(sys.executable)}" > nul
echo.
echo 기존 파일을 새 파일로 교체합니다...
xcopy "{new_program_folder_path}" "{application_path}" /E /H /C /I /Y > nul
echo.
echo 임시 업데이트 파일을 삭제합니다...
rmdir /s /q "{temp_update_folder}"
echo.
echo ========================================
echo  업데이트 완료!
echo ========================================
echo.
echo 3초 후에 프로그램을 다시 시작합니다.
timeout /t 3 /nobreak > nul
start "" "{os.path.join(application_path, os.path.basename(sys.executable))}"
del "%~f0"
            """)
        logging.info("Updater batch file created.")
        subprocess.Popen(updater_script_path, creationflags=subprocess.CREATE_NEW_CONSOLE)
        sys.exit(0)
    except Exception as e:
        logging.error(f"Update application failed: {e}")
        root_alert = tk.Tk(); root_alert.withdraw()
        messagebox.showerror("업데이트 실패", f"업데이트 적용 중 오류가 발생했습니다.\n\n{e}\n\n프로그램을 다시 시작해주세요.", parent=root_alert)
        root_alert.destroy()

def run_updater(repo_owner: str, repo_name: str, current_version: str):
    def check_thread():
        download_url, new_version = check_for_updates(repo_owner, repo_name, current_version)
        if download_url:
            root_alert = tk.Tk(); root_alert.withdraw()
            if messagebox.askyesno("업데이트 발견", f"새로운 버전({new_version})이 발견되었습니다.\n지금 업데이트하시겠습니까? (현재: {current_version})", parent=root_alert):
                root_alert.destroy()
                download_and_apply_update(download_url)
            else:
                root_alert.destroy(); logging.info("User declined the update.")
        else: logging.info("No new updates found.")
    threading.Thread(target=check_thread, daemon=True).start()

# ===================================================================
# 프로그램 본체
# ===================================================================
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ConfigManager:
    def __init__(self, config_path='config.xlsx'):
        self.config_path = config_path
        self.config = {
            'PALLET_SIZE': 60,
            'LEAD_TIME_DAYS': 2,
            'PALLETS_PER_TRUCK': 36,
            'MAX_TRUCKS_PER_DAY': 2,
            'FONT_SIZE': 11,
            'DELIVERY_DAYS': {str(i): 'True' if i < 5 else 'False' for i in range(7)},
            'NON_SHIPPING_DATES': [],
            'DAILY_TRUCK_OVERRIDES': {}
        }
        if os.path.exists(self.config_path):
            self.load_config()

    def load_config(self):
        try:
            settings_df = pd.read_excel(self.config_path, sheet_name='Settings').set_index('Setting')['Value']
            self.config['PALLET_SIZE'] = int(settings_df.get('PALLET_SIZE', 60))
            self.config['LEAD_TIME_DAYS'] = int(settings_df.get('LEAD_TIME_DAYS', 2))
            self.config['PALLETS_PER_TRUCK'] = int(settings_df.get('PALLETS_PER_TRUCK', 36))
            self.config['MAX_TRUCKS_PER_DAY'] = int(settings_df.get('MAX_TRUCKS_PER_DAY', 2))
            self.config['FONT_SIZE'] = int(settings_df.get('FONT_SIZE', 11))
            
            try:
                delivery_df = pd.read_excel(self.config_path, sheet_name='DeliveryConfig')
                self.config['DELIVERY_DAYS'] = delivery_df[~delivery_df['Key'].str.contains('NonShipping')].set_index('Key')['Value'].to_dict()
                non_shipping_dates = delivery_df[delivery_df['Key'] == 'NonShippingDates']['Value'].tolist()
                self.config['NON_SHIPPING_DATES'] = [datetime.datetime.strptime(str(d).split()[0], '%Y-%m-%d').date() for d in non_shipping_dates if d and isinstance(d, (str, datetime.datetime))]
            except Exception:
                logging.warning("DeliveryConfig 시트를 찾을 수 없거나 로드 오류가 발생했습니다. 기본값을 사용합니다.")

            try:
                truck_overrides_df = pd.read_excel(self.config_path, sheet_name='DailyTruckConfig')
                truck_overrides_df['Date'] = pd.to_datetime(truck_overrides_df['Date']).dt.date
                self.config['DAILY_TRUCK_OVERRIDES'] = truck_overrides_df.set_index('Date')['MaxTrucks'].to_dict()
                logging.info(f"{len(self.config['DAILY_TRUCK_OVERRIDES'])}개의 일자별 최대 차수 설정을 로드했습니다.")
            except Exception:
                logging.warning("DailyTruckConfig 시트를 찾을 수 없거나 로드 오류가 발생했습니다. 기본값을 사용합니다.")
                self.config['DAILY_TRUCK_OVERRIDES'] = {}

            logging.info("Config.xlsx 파일에서 설정을 성공적으로 로드했습니다.")
        except Exception as e:
            logging.error(f"Config load error: {e}")
            raise ValueError(f"`{self.config_path}` 로드 중 오류 발생: {e}. 시트 이름을 확인하세요.")

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
                
                delivery_data = [{'Key': key, 'Value': value} for key, value in config_data['DELIVERY_DAYS'].items()]
                delivery_data.extend([{'Key': 'NonShippingDates', 'Value': date.strftime('%Y-%m-%d')} for date in config_data['NON_SHIPPING_DATES']])
                delivery_df = pd.DataFrame(delivery_data)
                delivery_df.to_excel(writer, sheet_name='DeliveryConfig', index=False)
                
                if config_data.get('DAILY_TRUCK_OVERRIDES'):
                    overrides_data = [{'Date': date.strftime('%Y-%m-%d'), 'MaxTrucks': trucks} for date, trucks in config_data['DAILY_TRUCK_OVERRIDES'].items()]
                    overrides_df = pd.DataFrame(overrides_data)
                    overrides_df.to_excel(writer, sheet_name='DailyTruckConfig', index=False)

            self.config = config_data
            logging.info("설정을 config.xlsx 파일에 성공적으로 저장했습니다.")
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
        self.fixed_shipments = []
        self.fixed_shipment_reqs = []
        self.item_master_df = None
        self.allowed_models = []
        self.highlight_models = []
        self.unmet_demand_log = [] # '계획 실패' 경고를 위한 로그
        self._load_item_master()

    def _load_item_master(self):
        try:
            self.item_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'assets', 'Item.csv')
            if not os.path.exists(self.item_path):
                raise FileNotFoundError("assets/Item.csv 파일을 찾을 수 없습니다.")
                
            self.item_master_df = pd.read_csv(self.item_path)

            if 'Priority' not in self.item_master_df.columns:
                self.item_master_df['Priority'] = range(1, len(self.item_master_df) + 1)
            
            if 'SafetyStock' not in self.item_master_df.columns:
                self.item_master_df['SafetyStock'] = 0
            else:
                self.item_master_df['SafetyStock'] = pd.to_numeric(self.item_master_df['SafetyStock'], errors='coerce').fillna(0).astype(int)

            self.item_master_df.sort_values(by='Priority', inplace=True)
            self.allowed_models = self.item_master_df['Item Code'].tolist()
            self.highlight_models = self.item_master_df[self.item_master_df['Spec'].str.contains('HMC', na=False)]['Item Code'].tolist()
            
            self.item_master_df.set_index('Item Code', inplace=True)
            logging.info(f"Item.csv 로드 성공. 허용된 모델 수: {len(self.allowed_models)}")

        except Exception as e:
            messagebox.showerror("품목 정보 로드 실패", f"Item.csv 파일 처리 중 오류가 발생했습니다: {e}")
            logging.critical(f"Item.csv 로드 실패: {e}")
            raise
    
    def save_item_master(self):
        try:
            df_to_save = self.item_master_df.reset_index()
            df_to_save.to_csv(self.item_path, index=False)
            logging.info(f"품목 정보(최소 재고 포함)를 {self.item_path}에 저장했습니다.")
        except Exception as e:
            messagebox.showerror("품목 정보 저장 실패", f"Item.csv 파일 저장 중 오류 발생: {e}")
            logging.error(f"Item.csv 저장 실패: {e}")

    def process_plan_file(self):
        logging.info(f"파일 로드 시도: {self.current_filepath}")
        try:
            df_raw = pd.read_excel(self.current_filepath, sheet_name='《HCO&DIS》', header=None, engine='openpyxl')
            logging.info("원시 데이터 로드 성공. 헤더 행 탐색...")
            
            header_series = df_raw[11].astype(str)
            found = header_series.str.lower().str.contains('cover glass assy', na=False)
            if not found.any():
                raise ValueError("헤더 'Cover glass Assy'를 찾을 수 없습니다.")
            header_row_index = found.idxmax()

            logging.info(f"헤더 행 발견: {header_row_index}")
            df = df_raw.iloc[header_row_index:].copy()
            df.columns = df.iloc[0]
            df = df.iloc[1:].rename(columns={df.columns[11]: 'Model'})
            
            self.date_cols = sorted([col for col in df.columns if isinstance(col, (datetime.datetime, pd.Timestamp))])
            if not self.date_cols:
                raise ValueError("파일에서 유효한 날짜 컬럼을 찾을 수 없습니다.")

            logging.info(f"유효한 날짜 컬럼 {len(self.date_cols)}개 발견. 모델 필터링 시작...")
            df_filtered = df[df['Model'].isin(self.allowed_models)].copy()
            
            df_filtered.loc[:, self.date_cols] = df_filtered.loc[:, self.date_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
            agg_df = df_filtered.groupby('Model')[self.date_cols].sum()
            reindexed_df = agg_df.reindex(self.allowed_models).fillna(0).astype(int)
            
            self.aggregated_plan_df = reindexed_df.copy()
            logging.info(f"최종 집계된 DataFrame 생성 (shape: {self.aggregated_plan_df.shape})")
            return True
        except Exception as e:
            logging.error(f"Plan file processing error: {e}")
            raise
    
    def load_inventory_from_text(self, text_data):
        logging.info("재고 데이터 텍스트 파싱 시작...")
        data = []
        lines = [line.strip() for line in text_data.strip().split('\n') if line.strip()]
        inventory_date = None

        for line in lines:
            date_match = re.search(r'(\d{1,2})/(\d{1,2})', line)
            if date_match:
                month = int(date_match.group(1))
                day = int(date_match.group(2))
                year = datetime.date.today().year
                inventory_date = datetime.date(year, month, day)
                break
        
        for line in lines:
            matches = re.findall(r'(AAA\d+)\s+.*?(\d{1,3}(?:,\d{3})*)', line)
            for match in matches:
                model, inventory_str = match
                inventory = int(inventory_str.replace(',', ''))
                data.append({'Model': model, 'Inventory': inventory})

        i = 0
        while i < len(lines) - 1:
            current_line = lines[i]
            next_line = lines[i+1]
            if current_line.startswith('AAA') and next_line.replace(',', '').isdigit():
                model = current_line
                inventory = int(next_line.replace(',', ''))
                if not any(d['Model'] == model for d in data):
                    data.append({'Model': model, 'Inventory': inventory})
                i += 2
            else:
                i += 1
        
        if not data:
            raise ValueError("유효한 재고 데이터를 찾을 수 없습니다. 형식을 확인하세요.")
        
        inventory_df_raw = pd.DataFrame(data).set_index('Model')
        self.inventory_df = inventory_df_raw[inventory_df_raw.index.isin(self.allowed_models)]
        self.inventory_date = inventory_date if inventory_date else datetime.date.today()
        logging.info(f"재고 데이터 파싱 완료. 모델 수: {len(self.inventory_df)}, 기준일: {self.inventory_date}")

    def load_inventory_from_file(self, file_path):
        logging.info(f"파일에서 재고 데이터 로드 시작: {file_path}")
        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path, header=None)
            elif file_path.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path, header=None)
            else:
                raise ValueError("지원하지 않는 파일 형식입니다. (CSV, XLSX, XLS)")

            if df.shape[1] < 2:
                raise ValueError("파일은 최소 2개의 열(모델, 수량)을 포함해야 합니다.")

            df.rename(columns={0: 'Model', 1: 'Inventory'}, inplace=True)
            df['Model'] = df['Model'].astype(str)
            df['Inventory'] = pd.to_numeric(df['Inventory'], errors='coerce').fillna(0)

            inventory_df_raw = df[df['Model'].str.startswith('AAA', na=False)]
            inventory_df_raw = inventory_df_raw[['Model', 'Inventory']].set_index('Model')

            self.inventory_df = inventory_df_raw[inventory_df_raw.index.isin(self.allowed_models)]
            self.inventory_date = datetime.date.today()
            logging.info(f"파일 재고 로드 완료. 모델 수: {len(self.inventory_df)}, 기준일: {self.inventory_date}")
        except Exception as e:
            logging.error(f"Inventory file loading error: {e}")
            raise

    def run_simulation(self, adjustments=None, fixed_shipments=None, fixed_shipment_reqs=None):
        logging.info("개선된 로직으로 시뮬레이션을 시작합니다...")
        self.adjustments = adjustments if adjustments else []
        self.fixed_shipments = fixed_shipments if fixed_shipments else []
        self.fixed_shipment_reqs = fixed_shipment_reqs if fixed_shipment_reqs else []
        self.unmet_demand_log = [] # 시뮬레이션 시작 시 로그 초기화

        if self.aggregated_plan_df is None: return

        plan_df = self.aggregated_plan_df.copy()
        if self.inventory_df is not None:
            plan_df = plan_df.join(self.inventory_df, how='left').fillna({'Inventory': 0})
        else:
            plan_df = plan_df.assign(Inventory=0)
        plan_df['Inventory'] = plan_df['Inventory'].astype(int)

        if self.inventory_date:
            simulation_dates = [d for d in self.date_cols if d.date() >= self.inventory_date]
        else:
            simulation_dates = self.date_cols[:]
        
        if not simulation_dates: raise ValueError("시뮬레이션할 유효한 날짜가 없습니다.")

        lead_time = self.config.get('LEAD_TIME_DAYS', 2)
        pallet_size = self.config.get('PALLET_SIZE', 60)
        pallets_per_truck = self.config.get('PALLETS_PER_TRUCK', 36)
        truck_capacity = pallets_per_truck * pallet_size
        safety_stock = self.item_master_df['SafetyStock']

        for adj in self.adjustments:
            adj_date_dt = pd.to_datetime(adj['date'])
            if adj['model'] in plan_df.index and adj_date_dt in plan_df.columns:
                if adj['type'] == '수요':
                    plan_df.loc[adj['model'], adj_date_dt] += adj['qty']
                elif adj['type'] == '재고':
                    plan_df.loc[adj['model'], 'Inventory'] += adj['qty']

        demand_df = plan_df[simulation_dates]
        rolling_demand = demand_df.T.rolling(window=lead_time + 1, min_periods=1).sum().T
        
        inventory_over_time = pd.DataFrame(index=plan_df.index, columns=simulation_dates, dtype=np.int64)
        shipments_by_truck = {}
        current_inventory = plan_df['Inventory'].copy().astype(np.int64)
        
        for date in simulation_dates:
            daily_max_trucks = self.config.get('DAILY_TRUCK_OVERRIDES', {}).get(date.date(), self.config.get('MAX_TRUCKS_PER_DAY', 2))
            production_today = demand_df[date]
            is_shipping_day = self.config['DELIVERY_DAYS'].get(str(date.weekday()), 'False') == 'True'
            is_non_shipping_date = date.date() in self.config['NON_SHIPPING_DATES']
            total_shipments_today = pd.Series(0, index=plan_df.index, dtype=np.int64)

            if is_shipping_day and not is_non_shipping_date:
                # TIER 1: 필수 출고량 (Must-Ship)
                required_for_lead_time = rolling_demand[date] - current_inventory + safety_stock
                short_term_days = 7
                end_short_term = date + pd.Timedelta(days=short_term_days)
                short_term_dates = [d for d in simulation_dates if date <= d < end_short_term]
                required_for_short_term = demand_df[short_term_dates].sum(axis=1) - current_inventory + safety_stock
                must_ship_demand = pd.concat([required_for_lead_time, required_for_short_term], axis=1).max(axis=1)
                must_ship_demand[must_ship_demand < 0] = 0

                # TIER 2: 선제적 출고 가능량 (Pull-Forward)
                long_term_days = 30
                end_long_term = date + pd.Timedelta(days=long_term_days)
                long_term_dates = [d for d in simulation_dates if end_short_term <= d < end_long_term]
                pull_forward_demand = demand_df[long_term_dates].sum(axis=1) if long_term_dates else pd.Series(0, index=plan_df.index)

                fixed_reqs_today = pd.Series(0, index=plan_df.index, dtype=np.int64)
                for req in [r for r in self.fixed_shipment_reqs if r['date'] == date.date()]:
                    if req['model'] in fixed_reqs_today.index:
                        fixed_reqs_today.loc[req['model']] += req['qty']
                
                must_ship_demand = pd.concat([must_ship_demand, fixed_reqs_today], axis=1).max(axis=1).astype(np.int64)

                # 트럭 적재 로직
                fixed_shipments_today = pd.DataFrame(0, index=plan_df.index, columns=range(1, daily_max_trucks + 1), dtype=np.int64)
                truck_capacity_remains = pd.Series(truck_capacity, index=range(1, daily_max_trucks + 1), dtype=np.int64)
                
                for fixed in [s for s in self.fixed_shipments if s['date'] == date.date()]:
                    if fixed['truck_num'] <= daily_max_trucks:
                        fixed_shipments_today.loc[fixed['model'], fixed['truck_num']] = fixed['qty']
                        truck_capacity_remains[fixed['truck_num']] -= fixed['qty']
                
                remaining_must_ship = (must_ship_demand - fixed_shipments_today.sum(axis=1)).clip(lower=0)
                shipped_against_must = fixed_shipments_today.sum(axis=1).clip(upper=must_ship_demand)
                shipped_against_pull = (fixed_shipments_today.sum(axis=1) - shipped_against_must)
                remaining_pull_forward = (pull_forward_demand - shipped_against_pull).clip(lower=0)

                priority_models = self.item_master_df.sort_values('Priority').index
                auto_shipments_today = pd.DataFrame(0, index=plan_df.index, columns=range(1, daily_max_trucks + 1), dtype=np.int64)
                
                for truck_num in range(1, daily_max_trucks + 1):
                    if truck_capacity_remains[truck_num] <= 0: continue

                    for model in priority_models:
                        if truck_capacity_remains[truck_num] < pallet_size: break
                        if remaining_must_ship.loc[model] > 0:
                            pallets_needed = math.ceil(remaining_must_ship.loc[model] / pallet_size)
                            max_pallets_in_truck = math.floor(truck_capacity_remains[truck_num] / pallet_size)
                            pallets_to_ship = min(pallets_needed, max_pallets_in_truck)
                            if pallets_to_ship > 0:
                                qty_to_ship = pallets_to_ship * pallet_size
                                auto_shipments_today.loc[model, truck_num] += qty_to_ship
                                truck_capacity_remains[truck_num] -= qty_to_ship
                                remaining_must_ship.loc[model] -= qty_to_ship
                    
                    for model in priority_models:
                        if truck_capacity_remains[truck_num] < pallet_size: break
                        if remaining_pull_forward.loc[model] > 0:
                            pallets_can_ship = math.floor(remaining_pull_forward.loc[model] / pallet_size)
                            max_pallets_in_truck = math.floor(truck_capacity_remains[truck_num] / pallet_size)
                            pallets_to_ship = min(pallets_can_ship, max_pallets_in_truck)
                            if pallets_to_ship > 0:
                                qty_to_ship = pallets_to_ship * pallet_size
                                auto_shipments_today.loc[model, truck_num] += qty_to_ship
                                truck_capacity_remains[truck_num] -= qty_to_ship
                                remaining_pull_forward.loc[model] -= qty_to_ship

                final_daily_shipments = fixed_shipments_today + auto_shipments_today
                total_shipments_today = final_daily_shipments.sum(axis=1)

                for truck_num in range(1, daily_max_trucks + 1):
                    if truck_num not in shipments_by_truck:
                        shipments_by_truck[truck_num] = pd.DataFrame(0, index=plan_df.index, columns=simulation_dates, dtype=np.int64)
                    shipments_by_truck[truck_num][date] = final_daily_shipments[truck_num]
                
                unmet_demand = (must_ship_demand - total_shipments_today).clip(lower=0)
                if unmet_demand.sum() > 0:
                    for model, qty in unmet_demand[unmet_demand > 0].items():
                        log_entry = {'date': date.date(), 'model': model, 'unmet_qty': int(qty)}
                        self.unmet_demand_log.append(log_entry)
                        logging.warning(f"계획 실패: {log_entry}")

            current_inventory = current_inventory + total_shipments_today - production_today
            inventory_over_time[date] = current_inventory

        result_df = plan_df.copy()
        for date in simulation_dates:
            date_str = date.strftime("%m%d")
            result_df[f'재고_{date_str}'] = inventory_over_time[date]
            max_trucks = self.config.get('DAILY_TRUCK_OVERRIDES', {}).get(date.date(), self.config.get('MAX_TRUCKS_PER_DAY', 2))
            for truck_num in range(1, max_trucks + 1):
                col_name = f'출고_{truck_num}차_{date_str}'
                if truck_num in shipments_by_truck and date in shipments_by_truck[truck_num].columns:
                    result_df[col_name] = shipments_by_truck[truck_num][date]
                else: 
                    result_df[col_name] = 0

        self.simulated_plan_df = result_df.fillna(0).astype(int)
        logging.info("시뮬레이션 완료.")

    def find_and_propose_fix(self, max_truck_limit=3):
        if self.simulated_plan_df is None:
            return False, None

        df = self.simulated_plan_df
        inventory_cols = sorted([c for c in df.columns if isinstance(c, str) and c.startswith('재고_')])
        first_shortage_info = None

        for inv_col in inventory_cols:
            date_str = inv_col.split('_')[1]
            year = self.date_cols[0].year
            try:
                current_date = datetime.datetime.strptime(f"{year}-{date_str}", "%Y-%m%d").date()
            except ValueError:
                current_date = datetime.datetime.strptime(f"{year+1}-{date_str}", "%Y-%m%d").date()

            sorted_models = self.item_master_df.index
            shortage_series = df.loc[sorted_models, inv_col] < self.item_master_df.loc[sorted_models, 'SafetyStock']
            
            if shortage_series.any():
                model_to_fix = shortage_series.idxmax()
                first_shortage_info = {"model": model_to_fix, "shortage_date": current_date}
                break
        
        if not first_shortage_info:
            return False, None

        shortage_date = first_shortage_info['shortage_date']
        
        candidate_days = []
        start_date = self.inventory_date or self.date_cols[0].date()
        check_date = shortage_date - datetime.timedelta(days=1)
        
        while check_date >= start_date:
            is_shipping_day = self.config.get('DELIVERY_DAYS', {}).get(str(check_date.weekday()), 'False') == 'True'
            is_non_shipping_date = check_date in self.config['NON_SHIPPING_DATES']
            if is_shipping_day and not is_non_shipping_date:
                current_trucks = self.config.get('DAILY_TRUCK_OVERRIDES', {}).get(check_date, self.config.get('MAX_TRUCKS_PER_DAY'))
                candidate_days.append({"date": check_date, "trucks": current_trucks})
            check_date -= datetime.timedelta(days=1)
            
        if not candidate_days:
            return False, {"error": f"{first_shortage_info['model']} 부족({shortage_date})을 해결할 이전 납품일 없음"}

        eligible_candidates = [day for day in candidate_days if day['trucks'] < max_truck_limit]

        if not eligible_candidates:
            return False, {"error": f"재고 부족을 해결할 수 없습니다. 모든 유효한 이전 납품일의 차수가 최대({max_truck_limit}회)입니다."}

        eligible_candidates.sort(key=lambda x: (x['trucks'], -x['date'].toordinal()))
        
        best_candidate = eligible_candidates[0]
        
        fix_details = {
            "model": first_shortage_info['model'],
            "shortage_date": shortage_date,
            "shipping_date": best_candidate['date']
        }
        
        return True, fix_details
        
# --- 위젯 클래스 ---

class SearchableComboBox(ctk.CTkFrame):
    def __init__(self, parent, values):
        super().__init__(parent, fg_color="transparent")
        self.values = sorted(values)
        self.current_value = ""

        self.entry = ctk.CTkEntry(self, placeholder_text="모델 검색 또는 선택...")
        self.entry.pack(fill="x")
        self.entry.bind("<KeyRelease>", self.on_key_release)
        
        self.listbox = Listbox(self, height=5)
        self.listbox.bind("<<ListboxSelect>>", self.on_listbox_select)
        
        self.entry.bind("<FocusOut>", self.hide_listbox)
        self.listbox.bind("<FocusOut>", self.hide_listbox)
        self.bind("<FocusOut>", self.hide_listbox)

    def on_key_release(self, event=None):
        search_term = self.entry.get().lower()
        if not search_term:
            self.listbox.pack_forget()
            return

        matching_values = [v for v in self.values if search_term in v.lower()]
        
        if matching_values:
            self.listbox.delete(0, END)
            for val in matching_values:
                self.listbox.insert(END, val)
            self.listbox.pack(fill="x", expand=True)
            self.listbox.lift()
        else:
            self.listbox.pack_forget()

    def on_listbox_select(self, event=None):
        if self.listbox.curselection():
            self.current_value = self.listbox.get(self.listbox.curselection())
            self.entry.delete(0, END)
            self.entry.insert(0, self.current_value)
            self.listbox.pack_forget()
            self.focus()
            
    def hide_listbox(self, event=None):
        self.after(200, lambda: self.listbox.pack_forget() if self.focus_get() != self.listbox else None)

    def get(self):
        return self.entry.get()

class AdjustmentDialog(ctk.CTkToplevel):
    def __init__(self, parent, models):
        super().__init__(parent)
        self.models = models
        self.adjustments = []
        self.result = None
        self.title("수동 조정 입력")
        self.geometry("600x500")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(input_frame, text="모델:").grid(row=0, column=0, padx=5, pady=5)
        self.model_combo = SearchableComboBox(input_frame, values=self.models)
        self.model_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(input_frame, text="날짜 (YYYY-MM-DD):").grid(row=1, column=0, padx=5, pady=5)
        self.date_entry = ctk.CTkEntry(input_frame, placeholder_text=datetime.date.today().strftime('%Y-%m-%d'))
        self.date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(input_frame, text="수량:").grid(row=2, column=0, padx=5, pady=5)
        self.qty_entry = ctk.CTkEntry(input_frame)
        self.qty_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        ctk.CTkLabel(input_frame, text="타입:").grid(row=3, column=0, padx=5, pady=5)
        self.type_combo = ctk.CTkComboBox(input_frame, values=['재고', '수요', '고정 출고'])
        self.type_combo.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        
        button_frame = ctk.CTkFrame(self)
        button_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5)
        ctk.CTkButton(button_frame, text="추가", command=self.add_adjustment).pack()
        
        self.listbox = Listbox(self, height=10)
        self.listbox.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        ok_cancel_frame = ctk.CTkFrame(self, fg_color="transparent")
        ok_cancel_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="e")
        ctk.CTkButton(ok_cancel_frame, text="확인", command=self.ok_event).pack(side="left", padx=10)
        ctk.CTkButton(ok_cancel_frame, text="취소", command=self.cancel_event, fg_color="gray").pack(side="left")

        self.transient(parent)
        self.grab_set()

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
        self.listbox.insert(END, f"{adj['type']} | {adj['date']}, {adj['model']}, {adj['qty']:,}")
        self.qty_entry.delete(0, END)
        logging.info(f"조정 항목 추가: {adj}")

    def ok_event(self):
        self.result = self.adjustments
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()


class InventoryInputDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("재고 데이터 입력")
        self.geometry("450x350")
        self.result = None
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        prompt_frame = ctk.CTkFrame(self, fg_color="transparent")
        prompt_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        ctk.CTkLabel(prompt_frame, text="재고 데이터를 붙여넣거나 파일을 불러오세요.").pack(side="left")
        ctk.CTkButton(prompt_frame, text="파일에서 불러오기", command=self.load_file).pack(side="right")

        self.textbox = ctk.CTkTextbox(self, width=430, height=200)
        self.textbox.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="e")
        ctk.CTkButton(button_frame, text="확인", command=self.ok_event).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="취소", command=self.cancel_event, fg_color="gray").pack(side="left", padx=5)
        
        self.transient(parent)
        self.grab_set()
        self.textbox.focus()

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="재고 파일 선택 (Excel, CSV)",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if file_path:
            self.result = ('file', file_path)
            self.destroy()

    def ok_event(self):
        pasted_text = self.textbox.get("1.0", "end-1c")
        if pasted_text:
            self.result = ('text', pasted_text)
            self.destroy()
        else:
            messagebox.showwarning("입력 오류", "데이터를 입력하거나 파일을 선택해주세요.", parent=self)

    def cancel_event(self):
        self.result = None
        self.destroy()

class HolidayDialog(ctk.CTkToplevel):
    def __init__(self, parent, non_shipping_dates):
        super().__init__(parent)
        self.title("휴무일/공휴일 설정")
        self.geometry("300x350")
        self.result = None
        self.non_shipping_dates = [d for d in non_shipping_dates if isinstance(d, datetime.date)]
        
        self.cal = Calendar(self, selectmode='day')
        self.cal.pack(padx=10, pady=10, fill='x', expand=True)
        
        for date in self.non_shipping_dates:
            self.cal.calevent_add(date, 'holiday', 'holiday')
        self.cal.tag_config('holiday', background='red', foreground='white')
        
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="추가/제거", command=self.toggle_date).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="확인", command=self.ok_event).pack(side="left", padx=5)
        self.transient(parent)
        self.grab_set()

    def toggle_date(self):
        selected_date = self.cal.selection_get()
        if not selected_date: return
        
        if selected_date in self.non_shipping_dates:
            self.non_shipping_dates.remove(selected_date)
            event_ids = self.cal.get_calevents(selected_date, 'holiday')
            for event_id in event_ids:
                self.cal.calevent_remove(event_id)
        else:
            self.non_shipping_dates.append(selected_date)
            self.cal.calevent_add(selected_date, 'holiday', 'holiday')
        logging.info(f"휴무일 설정 변경: {self.non_shipping_dates}")

    def ok_event(self):
        self.result = self.non_shipping_dates
        self.destroy()

class DailyTruckDialog(ctk.CTkToplevel):
    def __init__(self, parent, overrides):
        super().__init__(parent)
        self.overrides = overrides.copy()
        self.result = None
        self.title("일자별 최대 차수 설정")
        self.geometry("500x450")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        ctk.CTkLabel(input_frame, text="날짜:").pack(side="left", padx=5)
        self.date_entry = DateEntry(input_frame, date_pattern='y-mm-dd', width=12)
        self.date_entry.pack(side="left", padx=5)

        ctk.CTkLabel(input_frame, text="최대 차수:").pack(side="left", padx=5)
        self.truck_entry = ctk.CTkEntry(input_frame, width=50)
        self.truck_entry.pack(side="left", padx=5)

        ctk.CTkButton(input_frame, text="추가/수정", command=self.add_override).pack(side="left", padx=10)

        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self.listbox = Listbox(list_frame, height=15)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        self.update_listbox()

        ctk.CTkButton(list_frame, text="선택 항목 삭제", command=self.remove_override).grid(row=1, column=0, pady=5)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=10, pady=10, sticky="e")
        ctk.CTkButton(button_frame, text="저장", command=self.ok_event).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="취소", command=self.cancel_event, fg_color="gray").pack(side="left")

        self.transient(parent)
        self.grab_set()

    def update_listbox(self):
        self.listbox.delete(0, END)
        sorted_overrides = sorted(self.overrides.items())
        for date, trucks in sorted_overrides:
            self.listbox.insert(END, f"{date.strftime('%Y-%m-%d')}  ->  {trucks}차")

    def add_override(self):
        try:
            date = self.date_entry.get_date()
            trucks = int(self.truck_entry.get())
            if trucks < 0:
                raise ValueError
            self.overrides[date] = trucks
            self.update_listbox()
        except (ValueError, TypeError):
            messagebox.showwarning("입력 오류", "유효한 날짜와 0 이상의 차수를 입력하세요.", parent=self)

    def remove_override(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        selected_text = self.listbox.get(selected_indices[0])
        date_str = selected_text.split(" ")[0]
        date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
        if date_obj in self.overrides:
            del self.overrides[date_obj]
        self.update_listbox()

    def ok_event(self):
        self.result = self.overrides
        self.destroy()

    def cancel_event(self):
        self.result = None
        self.destroy()

class SafetyStockDialog(ctk.CTkToplevel):
    def __init__(self, parent, item_master_df):
        super().__init__(parent)
        self.title("품목별 최소 재고 설정")
        self.geometry("500x600")
        self.result = None
        self.item_master_df = item_master_df.copy()
        self.entries = {}

        search_frame = ctk.CTkFrame(self)
        search_frame.pack(fill='x', padx=10, pady=5)
        ctk.CTkLabel(search_frame, text="품목 검색:").pack(side='left')
        self.search_entry = ctk.CTkEntry(search_frame)
        self.search_entry.pack(side='left', fill='x', expand=True, padx=5)
        self.search_entry.bind('<KeyRelease>', self.filter_items)

        header_frame = ctk.CTkFrame(self, fg_color="gray20")
        header_frame.pack(fill='x', padx=10, pady=(5,0))
        ctk.CTkLabel(header_frame, text="품목 코드", anchor='w', text_color="white").pack(side='left', expand=True, fill='x', padx=5)
        ctk.CTkLabel(header_frame, text="최소 재고 수량", anchor='e', text_color="white").pack(side='right', padx=20)

        self.scrollable_frame = ctk.CTkScrollableFrame(self)
        self.scrollable_frame.pack(expand=True, fill='both', padx=10, pady=(0,10))
        self.item_widgets = {}
        self.populate_items()
        
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(fill='x', padx=10, pady=10)
        ctk.CTkButton(button_frame, text="전체 0으로 설정", command=self.set_all_zero, fg_color="gray").pack(side='left', padx=10)
        ctk.CTkButton(button_frame, text="저장", command=self.save_and_close).pack(side='right', padx=10)
        ctk.CTkButton(button_frame, text="취소", command=self.cancel, fg_color="gray").pack(side='right')
        
        self.transient(parent)
        self.grab_set()

    def populate_items(self, filter_text=""):
        for frame in self.item_widgets.values():
            frame.destroy()
        self.item_widgets.clear()
        self.entries.clear()

        df_to_show = self.item_master_df
        if filter_text:
            df_to_show = df_to_show[df_to_show.index.str.contains(filter_text, case=False)]

        for item_code, row in df_to_show.iterrows():
            frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
            frame.pack(fill='x', pady=2)
            self.item_widgets[item_code] = frame

            label = ctk.CTkLabel(frame, text=item_code, anchor='w')
            label.pack(side='left', padx=5)
            
            entry = ctk.CTkEntry(frame, width=100, justify='right')
            entry.insert(0, str(row['SafetyStock']))
            entry.pack(side='right', padx=5)
            self.entries[item_code] = entry

    def filter_items(self, event=None):
        filter_text = self.search_entry.get()
        self.populate_items(filter_text)

    def set_all_zero(self):
        for entry in self.entries.values():
            entry.delete(0, END)
            entry.insert(0, "0")

    def save_and_close(self):
        try:
            for item_code, entry in self.entries.items():
                value = int(entry.get())
                self.item_master_df.loc[item_code, 'SafetyStock'] = value
            self.result = self.item_master_df
            self.destroy()
        except ValueError:
            messagebox.showerror("입력 오류", "최소 재고는 숫자로만 입력해야 합니다.", parent=self)

    def cancel(self):
        self.result = None
        self.destroy()

class ProductionPlannerApp(ctk.CTk):
    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.processor = PlanProcessor(self.config_manager.config)
        
        self.current_step = 0
        self.current_file = "파일이 로드되지 않았습니다."
        self.base_font_size = self.config_manager.config.get('FONT_SIZE', 11)
        self.font_big_bold = ctk.CTkFont(size=20, weight="bold")
        self.font_normal = ctk.CTkFont(size=self.base_font_size)
        self.font_small = ctk.CTkFont(size=self.base_font_size - 1)
        self.font_bold = ctk.CTkFont(size=self.base_font_size, weight="bold")
        self.font_italic = ctk.CTkFont(size=self.base_font_size, slant="italic")
        self.font_kpi = ctk.CTkFont(size=14, weight="bold")
        self.font_header = ctk.CTkFont(size=self.base_font_size + 1, weight="bold")
        self.font_edit = ctk.CTkFont(size=self.base_font_size, weight="bold")
        
        self.title(f"PlanForge Pro - 출고계획 시스템 ({CURRENT_VERSION})")
        self.geometry("1800x1000")
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")
        
        self.is_task_running = False
        self.thread_queue = Queue()

        self.sidebar_visible = True
        self.inventory_text_backup = None
        self.last_selected_model = None
        
        self.create_widgets()
        self.update_status_bar()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind_all("<Control-MouseWheel>", self.on_mouse_wheel_zoom)

        run_updater(REPO_OWNER, REPO_NAME, CURRENT_VERSION)
        
        self.after(100, self.process_thread_queue)
    
    def process_thread_queue(self):
        try:
            while True:
                task_name, data = self.thread_queue.get_nowait()
                
                if "update_ui" in task_name:
                    if task_name == "update_ui_step1":
                        self.update_ui_after_step1(data)
                    elif task_name == "update_ui_step2":
                        self.update_ui_after_step2(data)
                    elif task_name == "update_ui_step3":
                        self.update_ui_after_step3(data)
                elif "recalculation_done" in task_name:
                    self.update_ui_after_recalculation(data)
                elif task_name == "export_done":
                    messagebox.showinfo("내보내기 성공", f"계획이 {data}로 저장되었습니다.")
                elif task_name == "error":
                    messagebox.showerror("작업 오류", data)
                
                self.set_ui_task_state(False)

        except Empty:
            pass
        finally:
            self.after(100, self.process_thread_queue)

    def run_in_thread(self, worker_func):
        if self.is_task_running:
            messagebox.showwarning("작업 중", "이미 다른 작업이 실행 중입니다.")
            return
            
        self.set_ui_task_state(True)
        thread = threading.Thread(target=worker_func, daemon=True)
        thread.start()

    def set_ui_task_state(self, is_running):
        self.is_task_running = is_running
        state = "disabled" if is_running else "normal"
        self.step1_button.configure(state=state)
        self.step2_button.configure(state=state if self.current_step >=1 else "disabled")
        self.step3_button.configure(state=state if self.current_step >=2 else "disabled")
        self.step4_button.configure(state=state if self.current_step >=1 else "disabled")
        self.stabilize_button.configure(state=state if self.current_step >=2 else "disabled")
        self.save_settings_button.configure(state=state)
        self.daily_truck_button.configure(state=state)
        self.non_shipping_button.configure(state=state)
        self.safety_stock_button.configure(state=state)
        if not is_running:
            self.update_status_bar()

    def on_closing(self):
        try:
            self.unbind_all("<Control-MouseWheel>")
            plt.close('all')
            if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
                self.destroy()
        except Exception as e:
            logging.error(f"Closing error: {e}")
            self.destroy()

    def set_font_size(self, size):
        size = max(8, min(24, size)) 
        self.base_font_size = size
        self.config_manager.config['FONT_SIZE'] = size

        self.font_big_bold.configure(size=size + 9)
        self.font_normal.configure(size=size)
        self.font_small.configure(size=size - 1)
        self.font_bold.configure(size=size)
        self.font_italic.configure(size=size)
        self.font_kpi.configure(size=size + 3)
        self.font_header.configure(size=size + 1)
        self.font_edit.configure(size=size)

        self.sidebar_title.configure(font=self.font_big_bold)
        self.step1_button.configure(font=self.font_normal)
        self.step2_button.configure(font=self.font_normal)
        self.step3_button.configure(font=self.font_normal)
        self.step4_button.configure(font=self.font_normal)
        self.stabilize_button.configure(font=self.font_normal)
        self.font_size_title_label.configure(font=self.font_normal)
        self.font_minus_button.configure(font=self.font_normal)
        self.font_size_label.configure(text=str(size), font=self.font_normal)
        self.font_plus_button.configure(font=self.font_normal)
        self.settings_title_label.configure(font=self.font_bold)
        
        for label in self.setting_labels:
            label.configure(font=self.font_normal)
        for entry in self.settings_entries.values():
            entry.configure(font=self.font_normal)
        for cb in self.day_checkboxes.values():
            cb.configure(font=self.font_normal)
            
        self.daily_truck_button.configure(font=self.font_normal)
        self.non_shipping_button.configure(font=self.font_normal)
        self.safety_stock_button.configure(font=self.font_normal)
        self.save_settings_button.configure(font=self.font_normal)
        self.search_label.configure(font=self.font_normal)
        self.search_entry.configure(font=self.font_normal)
        self.lbl_models_found.configure(font=self.font_kpi)
        self.lbl_total_quantity.configure(font=self.font_kpi)
        self.lbl_date_range.configure(font=self.font_kpi)
        self.detail_tab_title.configure(font=self.font_bold)
        self.status_bar.configure(font=self.font_normal)

        if self.current_step > 0:
            self.filter_grid()

        logging.info(f"폰트 크기를 {size}로 변경했습니다.")

    def change_font_size(self, delta):
        self.set_font_size(self.base_font_size + delta)

    def prompt_for_font_size(self, event=None):
        dialog = ctk.CTkInputDialog(text="새 폰트 크기를 입력하세요:", title="폰트 크기 변경")
        new_size_str = dialog.get_input()
        if new_size_str:
            try:
                new_size = int(new_size_str)
                self.set_font_size(new_size)
            except (ValueError, TypeError):
                messagebox.showerror("입력 오류", "유효한 숫자를 입력해주세요.", parent=self)

    def on_mouse_wheel_zoom(self, event):
        self.set_font_size(self.base_font_size + (1 if event.delta > 0 else -1))

    def toggle_sidebar(self):
        if self.sidebar_visible:
            self.paned_window.remove(self.sidebar_frame)
            self.sidebar_toggle_button.configure(text="▶")
            self.sidebar_visible = False
        else:
            self.paned_window.insert(0, self.sidebar_frame)
            self.paned_window.sash_place(0, 280, 0)
            self.sidebar_toggle_button.configure(text="◀")
            self.sidebar_visible = True
    
    def create_widgets(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.paned_window = PanedWindow(self, orient=HORIZONTAL, sashrelief=tk.RAISED, bg="#D3D3D3")
        self.paned_window.grid(row=0, column=0, sticky="nsew")

        self.sidebar_frame = ctk.CTkFrame(self.paned_window, width=280, corner_radius=0)
        self.sidebar_frame.grid_rowconfigure(6, weight=1)
        self.paned_window.add(self.sidebar_frame, width=280)
        self.paned_window.paneconfigure(self.sidebar_frame, minsize=280)

        main_content_container = ctk.CTkFrame(self.paned_window, fg_color="transparent")
        main_content_container.grid_rowconfigure(0, weight=1)
        main_content_container.grid_columnconfigure(1, weight=1)
        self.paned_window.add(main_content_container)

        self.sidebar_toggle_button = ctk.CTkButton(main_content_container, text="◀", command=self.toggle_sidebar, width=20, height=40, corner_radius=5)
        self.sidebar_toggle_button.grid(row=0, column=0, sticky="w", pady=10)
        
        main_area_frame = ctk.CTkFrame(main_content_container, fg_color="transparent")
        main_area_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        main_area_frame.grid_columnconfigure(0, weight=1)
        main_area_frame.grid_rowconfigure(3, weight=1) # Row index changed for new frame

        self.sidebar_title = ctk.CTkLabel(self.sidebar_frame, text="PlanForge Pro", font=self.font_big_bold)
        self.sidebar_title.pack(pady=20)
        self.step1_button = ctk.CTkButton(self.sidebar_frame, text="1. 생산계획 불러오기", command=self.run_step1_aggregate, font=self.font_normal)
        self.step1_button.pack(fill='x', padx=20, pady=5)
        self.step2_button = ctk.CTkButton(self.sidebar_frame, text="2. 재고 반영 및 계획 시뮬레이션", command=self.run_step2_simulation, state="disabled", font=self.font_normal)
        self.step2_button.pack(fill='x', padx=20, pady=5)
        self.step3_button = ctk.CTkButton(self.sidebar_frame, text="3. 수동 조정 적용", command=self.run_step3_adjustments, state="disabled", font=self.font_normal)
        self.step3_button.pack(fill='x', padx=20, pady=5)
        self.step4_button = ctk.CTkButton(self.sidebar_frame, text="4. 계획 내보내기 (Excel)", command=self.export_to_excel, state="disabled", font=self.font_normal)
        self.step4_button.pack(fill='x', padx=20, pady=5)
        
        self.stabilize_button = ctk.CTkButton(self.sidebar_frame, text="✨ 재고 안정화 실행", command=self.run_stabilization, state="disabled", font=self.font_normal, fg_color="#0B5345", hover_color="#117A65")
        self.stabilize_button.pack(fill='x', padx=20, pady=(15, 5))

        font_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
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
        
        settings_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        settings_frame.pack(fill='x', expand=True, padx=20, pady=20)
        self.settings_title_label = ctk.CTkLabel(settings_frame, text="시스템 설정", font=self.font_bold)
        self.settings_title_label.pack()
        self.settings_entries = {}
        settings_map = {'팔레트당 수량': 'PALLET_SIZE', '리드타임 (일)': 'LEAD_TIME_DAYS', '트럭당 팔레트 수': 'PALLETS_PER_TRUCK', '기본 최대 차수': 'MAX_TRUCKS_PER_DAY'}
        self.setting_labels = []
        for label_text, key in settings_map.items():
            frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
            frame.pack(fill='x', pady=2)
            label = ctk.CTkLabel(frame, text=label_text, width=120, anchor='w', font=self.font_normal)
            label.pack(side='left')
            self.setting_labels.append(label)
            entry = ctk.CTkEntry(frame, font=self.font_normal)
            entry.pack(side='left', fill='x', expand=True)
            self.settings_entries[key] = entry
        
        self.delivery_days_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        self.delivery_days_frame.pack(fill='x', pady=5)
        ctk.CTkLabel(self.delivery_days_frame, text="납품 요일:", font=self.font_normal).pack(anchor="w", pady=(0, 2))
        
        checkbox_container = ctk.CTkFrame(self.delivery_days_frame, fg_color="transparent")
        checkbox_container.pack(fill='x')

        self.day_checkboxes = {}
        day_names = ["월", "화", "수", "목", "금", "토", "일"]
        for i, day in enumerate(day_names):
            state = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(i), 'False') == 'True'
            cb = ctk.CTkCheckBox(checkbox_container, text=day, onvalue=True, offvalue=False, font=self.font_normal, width=1)
            cb.grid(row=i // 4, column=i % 4, padx=2, pady=1, sticky='w')
            if state: cb.select()
            self.day_checkboxes[i] = cb

        self.daily_truck_button = ctk.CTkButton(settings_frame, text="일자별 최대 차수 설정", command=self.open_daily_truck_dialog, font=self.font_normal)
        self.daily_truck_button.pack(fill='x', padx=5, pady=5)
        self.non_shipping_button = ctk.CTkButton(settings_frame, text="휴무일/공휴일 설정", command=self.open_holiday_dialog, font=self.font_normal)
        self.non_shipping_button.pack(fill='x', padx=5, pady=5)
        
        self.safety_stock_button = ctk.CTkButton(settings_frame, text="품목별 최소 재고 설정", command=self.open_safety_stock_dialog, font=self.font_normal)
        self.safety_stock_button.pack(fill='x', padx=5, pady=5)

        self.save_settings_button = ctk.CTkButton(self.sidebar_frame, text="설정 저장 및 재계산", command=self.save_settings_and_recalculate, fg_color="#1F6AA5", font=self.font_normal)
        self.save_settings_button.pack(fill='x', padx=20, pady=10, side='bottom')
        self.load_settings_to_gui()

        top_frame = ctk.CTkFrame(main_area_frame, fg_color="transparent")
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.grid_columnconfigure(0, weight=1)
        search_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        search_frame.pack(fill='x', expand=True, pady=(0,5))
        self.search_label = ctk.CTkLabel(search_frame, text="품목 검색:", font=self.font_normal)
        self.search_label.pack(side='left', padx=(0,5))
        self.search_entry = ctk.CTkEntry(search_frame, font=self.font_normal)
        self.search_entry.pack(side='left', fill='x', expand=True)
        self.search_entry.bind("<KeyRelease>", self.filter_grid)
        self.kpi_frame = ctk.CTkFrame(top_frame, fg_color="#EAECEE", corner_radius=5)
        self.kpi_frame.pack(fill='x', expand=True)
        self.kpi_frame.grid_columnconfigure((0,1,2), weight=1)
        self.lbl_models_found = ctk.CTkLabel(self.kpi_frame, text="처리된 모델 수: -", font=self.font_kpi)
        self.lbl_models_found.grid(row=0, column=0, padx=10, pady=10)
        self.lbl_total_quantity = ctk.CTkLabel(self.kpi_frame, text="총생산량: -", font=self.font_kpi)
        self.lbl_total_quantity.grid(row=0, column=1, padx=10, pady=10)
        self.lbl_date_range = ctk.CTkLabel(self.kpi_frame, text="계획 기간: -", font=self.font_kpi)
        self.lbl_date_range.grid(row=0, column=2, padx=10, pady=10)
        
        # --- 경고 프레임 영역 수정 ---
        self.unmet_demand_frame = ctk.CTkFrame(main_area_frame, fg_color="#FFDDE1")
        self.unmet_demand_frame.grid(row=1, column=0, sticky="ew", pady=(5,0))
        self.unmet_demand_frame.grid_remove()
        unmet_title = ctk.CTkLabel(self.unmet_demand_frame, text="🚨 계획 실패 경고 (필수 출고량 부족)", font=self.font_bold, text_color="#C0392B")
        unmet_title.pack(pady=(5,0))
        self.unmet_list_frame = ctk.CTkScrollableFrame(self.unmet_demand_frame, label_text="", height=80)
        self.unmet_list_frame.pack(fill="x", expand=True, padx=5, pady=5)

        self.shortage_frame = ctk.CTkFrame(main_area_frame, fg_color="#FFF5E1")
        self.shortage_frame.grid(row=2, column=0, sticky="ew", pady=5)
        self.shortage_frame.grid_remove()
        shortage_title = ctk.CTkLabel(self.shortage_frame, text="⚠️ 재고 부족 경고 (안전 재고 미달)", font=self.font_bold, text_color="#E67E22")
        shortage_title.pack(pady=(5,0))
        self.shortage_list_frame = ctk.CTkScrollableFrame(self.shortage_frame, label_text="", height=80)
        self.shortage_list_frame.pack(fill="x", expand=True, padx=5, pady=5)
        
        self.tabview = ctk.CTkTabview(main_area_frame)
        self.tabview.grid(row=3, column=0, sticky="nsew") # Row index changed
        self.master_tab = self.tabview.add("개요")
        self.detail_tab = self.tabview.add("상세")
        self.master_tab.grid_columnconfigure(0, weight=1)
        self.master_tab.grid_rowconfigure(0, weight=1)
        self.detail_tab.grid_columnconfigure(0, weight=1)
        self.detail_tab.grid_rowconfigure(1, weight=1)

        master_container = ctk.CTkFrame(self.master_tab, fg_color="transparent")
        master_container.grid(row=0, column=0, sticky="nsew")
        master_container.grid_rowconfigure(0, weight=1)
        master_container.grid_columnconfigure(0, weight=1)
        canvas = tk.Canvas(master_container, highlightthickness=0, bg="#f2f2f2")
        v_scrollbar = ctk.CTkScrollbar(master_container, orientation="vertical", command=canvas.yview)
        h_scrollbar = ctk.CTkScrollbar(master_container, orientation="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.master_frame = ctk.CTkFrame(canvas, fg_color="transparent")
        canvas.create_window((0, 0), window=self.master_frame, anchor="nw")
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.master_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        self.detail_tab_title = ctk.CTkLabel(self.detail_tab, text="상세: 선택된 모델의 출고 시뮬레이션", font=self.font_bold)
        self.detail_tab_title.grid(row=0, column=0, sticky="w", padx=10, pady=(5,0))
        self.detail_frame = ctk.CTkScrollableFrame(self.detail_tab, label_text="")
        self.detail_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        self.status_bar = ctk.CTkLabel(self, text="준비 완료", anchor="w", font=self.font_normal)
        self.status_bar.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))

    def run_step1_aggregate(self):
        file_path = filedialog.askopenfilename(title="생산계획 엑셀 파일 선택", filetypes=(("Excel", "*.xlsx *.xls"),))
        if not file_path: return
        self.update_status_bar("생산계획을 집계하는 중입니다...")
        
        def worker():
            try:
                self.processor.current_filepath = file_path
                self.processor.process_plan_file()
                self.current_file = os.path.basename(file_path)
                df = self.processor.aggregated_plan_df
                if df is None or df.empty:
                    self.thread_queue.put(("error", "처리할 생산 계획 데이터가 없습니다."))
                    return
                plan_cols = self.processor.date_cols
                df_filtered = df[df[plan_cols].sum(axis=1) > 0]
                result = { "models_found": len(df_filtered.index), "total_qty": df_filtered[plan_cols].sum().sum(), "date_range": f"{plan_cols[0].strftime('%y/%m/%d')} ~ {plan_cols[-1].strftime('%y/%m/%d')}" }
                self.thread_queue.put(("update_ui_step1", result))
            except Exception as e:
                self.thread_queue.put(("error", f"1단계 파일 처리 실패: {e}"))
        self.run_in_thread(worker)

    def update_ui_after_step1(self, data):
        self.current_step = 1
        self.lbl_models_found.configure(text=f"처리된 모델 수: {data['models_found']} 개")
        self.lbl_total_quantity.configure(text=f"총생산량: {data['total_qty']:,.0f} 개")
        self.lbl_date_range.configure(text=f"계획 기간: {data['date_range']}")
        [widget.destroy() for widget in self.detail_frame.winfo_children()]
        self.filter_grid()
        self.update_status_bar("1단계: 생산계획 집계 완료")
        self.shortage_frame.grid_remove()
        self.unmet_demand_frame.grid_remove()
        logging.info("1단계 완료. UI 업데이트 완료.")

    def run_step2_simulation(self):
        dialog = InventoryInputDialog(self)
        self.wait_window(dialog)
        result = dialog.result
        if not result: return
        self.update_status_bar("출고 계획을 시뮬레이션하는 중입니다...")
        source_type, param = result
        def worker():
            try:
                if source_type == 'text':
                    self.processor.load_inventory_from_text(param)
                    self.inventory_text_backup = param
                elif source_type == 'file':
                    self.processor.load_inventory_from_file(param)
                    self.inventory_text_backup = None
                self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments, fixed_shipment_reqs=self.processor.fixed_shipment_reqs)
                if self.processor.simulated_plan_df is None:
                    self.thread_queue.put(("error", "시뮬레이션 결과가 생성되지 않았습니다."))
                    return
                total_ship = self.processor.simulated_plan_df[[c for c in self.processor.simulated_plan_df.columns if isinstance(c, str) and c.startswith('출고_')]].sum().sum()
                self.thread_queue.put(("update_ui_step2", {"total_ship": total_ship}))
            except Exception as e:
                self.thread_queue.put(("error", f"2단계 시뮬레이션 실패: {e}"))
        self.run_in_thread(worker)

    def update_ui_after_step2(self, data):
        self.current_step = 2
        self.lbl_total_quantity.configure(text=f"총출고량: {data['total_ship']:,.0f} 개")
        [widget.destroy() for widget in self.detail_frame.winfo_children()]
        self.filter_grid()
        self.update_status_bar("2단계: 출고 계획 시뮬레이션 완료.")
        self.check_shipment_capacity()
        self.update_unmet_demand_warnings() # 계획 실패 경고 업데이트
        self.update_shortage_warnings()
        logging.info("2단계 완료. 시뮬레이션 결과 UI 업데이트 완료.")

    def run_step3_adjustments(self):
        dialog = AdjustmentDialog(self, models=self.processor.allowed_models)
        self.wait_window(dialog)
        all_adjustments = dialog.result
        if all_adjustments is None: return
        self.update_status_bar("수동 조정을 적용하여 재계산 중입니다...")
        
        def worker():
            try:
                self.processor.adjustments = [adj for adj in all_adjustments if adj['type'] in ['재고', '수요']]
                self.processor.fixed_shipment_reqs = [adj for adj in all_adjustments if adj['type'] == '고정 출고']
                if self.inventory_text_backup:
                    self.processor.load_inventory_from_text(self.inventory_text_backup)
                self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments, fixed_shipment_reqs=self.processor.fixed_shipment_reqs)
                total_ship = self.processor.simulated_plan_df[[c for c in self.processor.simulated_plan_df.columns if isinstance(c, str) and c.startswith('출고_')]].sum().sum()
                self.thread_queue.put(("update_ui_step3", {"total_ship": total_ship}))
            except Exception as e:
                self.thread_queue.put(("error", f"3단계 조정 실패: {e}"))
        self.run_in_thread(worker)
        
    def update_ui_after_step3(self, data):
        self.current_step = 3
        self.lbl_total_quantity.configure(text=f"총출고량: {data['total_ship']:,.0f} 개")
        [widget.destroy() for widget in self.detail_frame.winfo_children()]
        self.filter_grid()
        self.update_status_bar("3단계: 수동 조정 적용 완료.")
        self.check_shipment_capacity()
        self.update_unmet_demand_warnings() # 계획 실패 경고 업데이트
        self.update_shortage_warnings()
        logging.info("3단계 완료. 조정 결과 UI 업데이트 완료.")
    
    def run_stabilization(self):
        if self.current_step < 2: return
        
        if not messagebox.askyesno("재고 안정화 실행", "자동으로 최적의 출고 차수를 계산합니다.\n\n이 작업은 여러 번의 재계산을 포함하며, '일자별 최대 차수 설정'이 변경될 수 있습니다.\n계속하시겠습니까?", parent=self):
            return

        self.update_status_bar("재고 안정화 실행 중... (0% 완료)")
        
        def worker():
            try:
                max_iterations = 30
                max_truck_limit_per_day = 3

                for i in range(max_iterations):
                    progress = int(((i + 1) / max_iterations) * 100)
                    self.status_bar.configure(text=f"재고 안정화 실행 중... ({progress}% 완료, {i+1}차 재계산)")

                    if self.inventory_text_backup: self.processor.load_inventory_from_text(self.inventory_text_backup)
                    self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments, fixed_shipment_reqs=self.processor.fixed_shipment_reqs)
                    
                    # 1순위: 계획 실패(필수 물량 부족) 해결
                    if self.processor.unmet_demand_log:
                        first_failure = self.processor.unmet_demand_log[0]
                        shipping_date = first_failure['date']
                        current_max = self.processor.config.get('DAILY_TRUCK_OVERRIDES', {}).get(shipping_date, self.processor.config.get('MAX_TRUCKS_PER_DAY'))
                        if current_max >= max_truck_limit_per_day:
                            self.thread_queue.put(("error", f"안정화 실패: {shipping_date}의 차수가 이미 최대({max_truck_limit_per_day}회)입니다."))
                            return
                        new_max = current_max + 1
                        self.processor.config['DAILY_TRUCK_OVERRIDES'][shipping_date] = new_max
                        logging.info(f"안정화 {i+1}단계: {shipping_date} 필수물량 부족 해결 위해 차수를 {new_max}로 증가")
                        continue # 다음 루프에서 재시도
                    
                    # 2순위: 안전 재고 부족 해결
                    found_shortage, fix_info = self.processor.find_and_propose_fix(max_truck_limit=max_truck_limit_per_day)
                    if not found_shortage:
                        msg = f"재고 안정화 완료! ({i}회 반복)" if i > 0 else "현재 계획은 안정적입니다."
                        self.thread_queue.put(("recalculation_done", msg))
                        return
                    
                    shipping_date = fix_info['shipping_date']
                    current_max = self.processor.config.get('DAILY_TRUCK_OVERRIDES', {}).get(shipping_date, self.processor.config.get('MAX_TRUCKS_PER_DAY'))
                    new_max = current_max + 1
                    self.processor.config['DAILY_TRUCK_OVERRIDES'][shipping_date] = new_max
                    logging.info(f"안정화 {i+1}단계: {fix_info['shortage_date']} 재고부족({fix_info['model']}) 해결 위해 {shipping_date} 차수를 {new_max}로 증가")

                self.thread_queue.put(("error", f"최적화 실패: 최대 반복 횟수({max_iterations}회) 초과"))

            except Exception as e:
                self.thread_queue.put(("error", f"재고 안정화 중 오류 발생: {e}"))
                
        self.run_in_thread(worker)

    def recalculate_with_fixed_values(self):
        self.update_status_bar("고정값을 적용하여 재계산 중입니다...")
        def worker():
            try:
                if self.inventory_text_backup:
                    self.processor.load_inventory_from_text(self.inventory_text_backup)
                self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments, fixed_shipment_reqs=self.processor.fixed_shipment_reqs)
                self.thread_queue.put(("recalculation_done", "재계산 완료"))
            except Exception as e:
                self.thread_queue.put(("error", f"재계산 실패: {e}"))
        self.run_in_thread(worker)

    def save_settings_and_recalculate(self):
        self.update_status_bar("설정 저장 및 재계산 중...")
        new_config = self.config_manager.config.copy()
        try:
            for key, entry_widget in self.settings_entries.items():
                new_config[key] = int(entry_widget.get())
            new_delivery_days = {str(i): str(self.day_checkboxes[i].get()) for i in range(7)}
            new_config['DELIVERY_DAYS'] = new_delivery_days
        except Exception as e:
            messagebox.showerror("설정 오류", f"설정값 저장 중 오류 발생: {e}"); return

        def worker():
            try:
                self.config_manager.save_config(new_config)
                self.processor.config = new_config
                if self.current_step >= 1: self.processor.process_plan_file()
                if self.current_step >= 2:
                    if self.inventory_text_backup: self.processor.load_inventory_from_text(self.inventory_text_backup)
                    self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments, fixed_shipment_reqs=self.processor.fixed_shipment_reqs)
                self.thread_queue.put(("recalculation_done", "설정 저장 및 재계산 완료"))
            except Exception as e:
                self.thread_queue.put(("error", f"재계산 실패: {e}"))
        self.run_in_thread(worker)

    def update_ui_after_recalculation(self, message):
        self.filter_grid()
        if self.last_selected_model:
            self.populate_detail_view(self.last_selected_model)
        self.update_unmet_demand_warnings()
        self.update_shortage_warnings(propose_fix=False)
        self.update_status_bar(message)
        messagebox.showinfo("성공", message)
        logging.info(message)
        
    def export_to_excel(self):
        if self.current_step < 1:
            messagebox.showwarning("오류", "먼저 데이터를 불러와야 합니다.")
            return
        start_date = self.processor.date_cols[0].strftime('%m-%d')
        end_date = self.processor.date_cols[-1].strftime('%m-%d')
        filename = f"{start_date}~{end_date} {'생산계획' if self.current_step == 1 else '출고계획'}.xlsx"
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=(("Excel", "*.xlsx"),))
        if not file_path: return
        self.update_status_bar(f"'{os.path.basename(file_path)}' 파일로 내보내는 중입니다...")
        
        def worker():
            try:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    if self.current_step == 1:
                        df = self.processor.aggregated_plan_df
                        df_filtered = df[df[self.processor.date_cols].sum(axis=1) > 0]
                        df_filtered.to_excel(writer, sheet_name='생산계획')
                    else:
                        df = self.processor.simulated_plan_df
                        shipment_cols = [c for c in df.columns if isinstance(c, str) and c.startswith('출고_')]
                        df_filtered = df[df[shipment_cols].sum(axis=1) > 0]
                        df_filtered.to_excel(writer, sheet_name='Full Plan')
                        all_truck_nums = sorted(list(set(int(c.split('_')[1][:-1]) for c in shipment_cols)))
                        for truck_num in all_truck_nums:
                            sheet_name = f'{truck_num}차 출고'
                            cols_for_truck = [f'출고_{truck_num}차_{d.strftime("%m%d")}' for d in self.processor.date_cols if f'출고_{truck_num}차_{d.strftime("%m%d")}' in df.columns]
                            if not cols_for_truck: continue
                            df_truck = df_filtered[cols_for_truck].copy()
                            df_truck.columns = [c[-4:] for c in cols_for_truck]
                            df_truck = df_truck.rename(columns=lambda x: f"{x[:2]}-{x[2:]}")
                            df_truck[df_truck.sum(axis=1) > 0].to_excel(writer, sheet_name=sheet_name)
                self.thread_queue.put(("export_done", file_path))
            except Exception as e:
                self.thread_queue.put(("error", f"내보내기 실패: {e}"))
        self.run_in_thread(worker)

    def filter_grid(self, event=None):
        df_source = self.processor.aggregated_plan_df if self.current_step == 1 else self.processor.simulated_plan_df
        if df_source is None:
            df_to_show = pd.DataFrame()
        else:
            sum_cols = []
            if self.current_step == 1 and self.processor.date_cols:
                sum_cols = self.processor.date_cols
            elif self.current_step >= 2:
                sum_cols = [c for c in df_source.columns if isinstance(c, str) and c.startswith('출고_')]
            df_to_show = df_source[df_source[sum_cols].sum(axis=1) > 0].copy() if sum_cols else df_source.copy()

        search_term = self.search_entry.get().lower()
        if search_term:
            df_to_show = df_to_show[df_to_show.index.str.lower().str.contains(search_term)]

        self.populate_master_grid_from_scratch(df_to_show)

    def populate_master_grid_from_scratch(self, df_to_show):
        for widget in self.master_frame.winfo_children():
            widget.destroy()
        if df_to_show.empty: return

        all_plan_cols = self.processor.date_cols
        display_cols = []
        for date_col in all_plan_cols:
            is_shipping_day = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(date_col.weekday()), 'False') == 'True'
            if is_shipping_day and date_col.date() not in self.config_manager.config.get('NON_SHIPPING_DATES', []):
                display_cols.append(date_col)

        if self.current_step < 2:
            df_display = df_to_show.reset_index()
            headers = ['Model'] + [d.strftime('%m-%d') for d in display_cols]
            for c, h_text in enumerate(headers):
                ctk.CTkLabel(self.master_frame, text=h_text, font=self.font_header).grid(row=0, column=c, sticky="ew", padx=1, pady=2)
            for r, row_data in df_display.iterrows():
                model = row_data['Model']
                bg = "#D6EAF8" if model in self.processor.highlight_models else "transparent"
                lbl_model = ctk.CTkLabel(self.master_frame, text=model, fg_color=bg, font=self.font_normal, anchor="w", padx=5)
                lbl_model.grid(row=r + 1, column=0, sticky="ew")
                lbl_model.bind("<Double-Button-1>", lambda e, m=model: self.on_row_double_click(m))
                for i, date_col in enumerate(display_cols):
                    val = row_data.get(date_col, 0)
                    lbl_data = ctk.CTkLabel(self.master_frame, text=f"{val:,.0f}", fg_color=bg, font=self.font_normal, anchor="e", padx=5)
                    lbl_data.grid(row=r + 1, column=i + 1, sticky="ew")
                    lbl_data.bind("<Double-Button-1>", lambda e, m=model: self.on_row_double_click(m))
        else:
            df_display = df_to_show.reset_index()
            active_trucks_per_day = {}
            for d in display_cols:
                date_str_md = d.strftime("%m%d")
                truck_cols_for_day = [str(c) for c in df_to_show.columns if str(c).startswith('출고_') and str(c).endswith(f"_{date_str_md}")]
                active_trucks = []
                if truck_cols_for_day:
                    truck_nums = [int(c.split('_')[1][:-1]) for c in truck_cols_for_day]
                    for num in sorted(list(set(truck_nums))):
                        col_name = f'출고_{num}차_{date_str_md}'
                        if col_name in df_to_show.columns and df_to_show[col_name].sum() > 0:
                            active_trucks.append(num)
                if active_trucks:
                    active_trucks_per_day[d.date()] = active_trucks
            
            ctk.CTkLabel(self.master_frame, text="Model", font=self.font_header).grid(row=0, column=0, rowspan=2, sticky="nsew", padx=1, pady=2)
            current_col_idx, col_idx_map = 1, {}
            for d in display_cols:
                date_obj = d.date()
                if date_obj in active_trucks_per_day:
                    trucks_to_display = active_trucks_per_day[date_obj]
                    num_trucks = len(trucks_to_display)
                    if num_trucks > 0:
                        ctk.CTkLabel(self.master_frame, text=f"{d.strftime('%m-%d')}", font=self.font_header, justify="center").grid(row=0, column=current_col_idx, columnspan=num_trucks, sticky="ew", padx=1, pady=2)
                        col_idx_map[date_obj] = (current_col_idx, trucks_to_display)
                        for truck_num in trucks_to_display:
                            ctk.CTkLabel(self.master_frame, text=f"{truck_num}차", font=self.font_header).grid(row=1, column=current_col_idx, sticky="ew", padx=1, pady=2)
                            current_col_idx += 1
            
            for r, row_data in df_display.iterrows():
                row_idx, model = r + 2, row_data['Model']
                bg_color = "#D6EAF8" if model in self.processor.highlight_models else "#FFFFFF"
                lbl_model = ctk.CTkLabel(self.master_frame, text=model, fg_color=bg_color, font=self.font_normal, anchor="w", padx=5)
                lbl_model.grid(row=row_idx, column=0, sticky="ew")
                lbl_model.bind("<Double-Button-1>", lambda e, m=model: self.on_row_double_click(m))
                for date_col in display_cols:
                    date_obj = date_col.date()
                    if date_obj in col_idx_map:
                        start_col, trucks_to_display = col_idx_map[date_obj]
                        for i, truck_num in enumerate(trucks_to_display):
                            col_name = f'출고_{truck_num}차_{date_col.strftime("%m%d")}'
                            val = row_data.get(col_name, 0)
                            is_fixed = any(s['model'] == model and s['date'] == date_col.date() and s['truck_num'] == truck_num for s in self.processor.fixed_shipments)
                            text, label_bg, text_color = f"{val:,.0f}" if val else "0", bg_color, "black"
                            if is_fixed: label_bg, text_color = "#A9CCE3", "blue"
                            data_label = ctk.CTkLabel(self.master_frame, text=text, fg_color=label_bg, font=self.font_bold if is_fixed else self.font_normal, anchor="e", padx=5, text_color=text_color)
                            data_label.grid(row=row_idx, column=start_col + i, sticky="ew")
                            data_label.bind("<Double-Button-1>", lambda e, m=model, d=date_col.date(), t=truck_num: self.on_shipment_double_click(e, m, d, t))
                            data_label.bind("<Button-3>", lambda e, m=model, d=date_col.date(), t=truck_num: self.on_shipment_right_click(e, m, d, t))

            summary_row_idx = len(df_display) + 2
            ctk.CTkLabel(self.master_frame, text="합계", fg_color="#EAECEE", font=self.font_bold, anchor="center").grid(row=summary_row_idx, column=0, sticky="ew", padx=1, pady=2)
            pallet_size = self.config_manager.config.get('PALLET_SIZE', 60)
            for date_col in display_cols:
                date_obj = date_col.date()
                if date_obj in col_idx_map:
                    start_col, trucks_to_display = col_idx_map[date_obj]
                    for i, truck_num in enumerate(trucks_to_display):
                        col_name = f'출고_{truck_num}차_{date_col.strftime("%m%d")}'
                        total_sum, total_pallets = 0, 0
                        if col_name in df_to_show.columns:
                            total_sum = df_to_show[col_name].sum()
                            total_pallets = df_to_show[col_name].apply(lambda qty: math.ceil(qty / pallet_size) if qty > 0 else 0).sum()
                        summary_text = f"{total_sum:,.0f}\n({int(total_pallets)} P)"
                        ctk.CTkLabel(self.master_frame, text=summary_text, fg_color="#EAECEE", font=self.font_bold, anchor="e", padx=5).grid(row=summary_row_idx, column=start_col + i, sticky="ew", padx=1, pady=2)

    def on_row_double_click(self, model_name):
        if self.current_step < 2: return
        self.last_selected_model = model_name
        self.populate_detail_view(model_name)
        self.tabview.set("상세")
        self.update_status_bar(f"'{model_name}'의 상세 정보를 표시합니다.")
    
    def on_shipment_double_click(self, event, model, date, truck_num):
        if self.current_step < 2: return
        dialog = ctk.CTkInputDialog(text=f"'{model}' {date.strftime('%m-%d')} {truck_num}차 출고량을 수동으로 수정/고정합니다.\n\n(현재값: {self.processor.simulated_plan_df.loc[model, f'출고_{truck_num}차_{date.strftime("%m%d")}']:,.0f})", title="출고량 수동 고정")
        new_value_str = dialog.get_input()
        if new_value_str is not None:
            try:
                new_value = int(new_value_str)
                if new_value < 0: raise ValueError
                self.fix_shipment(model, date, new_value, truck_num)
                self.recalculate_with_fixed_values()
            except (ValueError, TypeError):
                messagebox.showerror("입력 오류", "유효한 숫자를 입력해주세요.", parent=self)

    def on_shipment_right_click(self, event, model, date, truck_num):
        if self.current_step < 2: return
        menu = Menu(self, tearoff=0)
        is_fixed = any(s['model'] == model and s['date'] == date and s['truck_num'] == truck_num for s in self.processor.fixed_shipments)
        if is_fixed:
            menu.add_command(label=f"{truck_num}차 고정 해제", command=lambda: self.unfix_shipment(model, date, truck_num))
        else:
            menu.add_command(label=f"{truck_num}차 고정", command=lambda: self.on_fix_request(model, date, truck_num))
        menu.tk_popup(event.x_root, event.y_root)

    def on_fix_request(self, model, date, truck_num):
        if self.processor.simulated_plan_df is None: return
        shipment_value = self.processor.simulated_plan_df.loc[model, f'출고_{truck_num}차_{date.strftime("%m%d")}']
        self.fix_shipment(model, date, shipment_value, truck_num)
        self.recalculate_with_fixed_values()

    def fix_shipment(self, model, date, qty, truck_num):
        self.processor.fixed_shipments = [s for s in self.processor.fixed_shipments if not (s['model'] == model and s['date'] == date and s['truck_num'] == truck_num)]
        self.processor.fixed_shipments.append({'model': model, 'date': date, 'qty': qty, 'truck_num': truck_num})
        logging.info(f"출고량 고정: {model}, {date}, {qty}, {truck_num}차")

    def unfix_shipment(self, model, date, truck_num):
        self.processor.fixed_shipments = [s for s in self.processor.fixed_shipments if not (s['model'] == model and s['date'] == date and s['truck_num'] == truck_num)]
        self.recalculate_with_fixed_values()
        logging.info(f"출고량 고정 해제: {model}, {date}, {truck_num}차")
        
    def update_unmet_demand_warnings(self):
        for widget in self.unmet_list_frame.winfo_children():
            widget.destroy()
        if self.processor.unmet_demand_log:
            self.unmet_demand_frame.grid()
            for log in self.processor.unmet_demand_log:
                msg = f"{log['date'].strftime('%m/%d')} {log['model']}: 필수 출고량 {log['unmet_qty']:,}개 부족"
                ctk.CTkLabel(self.unmet_list_frame, text=msg, font=self.font_small, anchor="w").pack(fill="x", padx=5)
        else:
            self.unmet_demand_frame.grid_remove()

    def update_shortage_warnings(self, propose_fix=True):
        for widget in self.shortage_list_frame.winfo_children():
            widget.destroy()
        df = self.processor.simulated_plan_df
        if df is None:
            self.shortage_frame.grid_remove()
            return

        inventory_cols = sorted([c for c in df.columns if isinstance(c, str) and c.startswith('재고_')])
        if not inventory_cols:
            self.shortage_frame.grid_remove()
            return
        
        shortages_info, shortage_messages = [], []
        sorted_models = self.processor.item_master_df.index
        for model in sorted_models:
            if model not in df.index: continue
            safety_stock = self.processor.item_master_df.loc[model, 'SafetyStock']
            for inv_col in inventory_cols:
                if df.loc[model, inv_col] < safety_stock:
                    date_str = inv_col.split('_')[1]
                    year = self.processor.date_cols[0].year
                    try: date_obj = datetime.datetime.strptime(f"{year}-{date_str}", "%Y-%m%d").date()
                    except ValueError: date_obj = datetime.datetime.strptime(f"{year+1}-{date_str}", "%Y-%m%d").date()
                    info = {"model": model, "shortage_date": date_obj, "current_stock": df.loc[model, inv_col], "safety_stock": safety_stock}
                    shortages_info.append(info)
                    shortage_messages.append(f"{model}: {date_obj.strftime('%m/%d')} 재고({info['current_stock']:,}) < 최소({info['safety_stock']:,})")
                    break
        
        if shortages_info:
            self.shortage_frame.grid()
            for msg in shortage_messages:
                ctk.CTkLabel(self.shortage_list_frame, text=msg, font=self.font_small, anchor="w").pack(fill="x", padx=5)
            if propose_fix and not self.is_task_running:
                self.propose_shortage_fix(shortages_info)
        else:
            self.shortage_frame.grid_remove()

    def propose_shortage_fix(self, shortages_info):
        first_shortage = sorted(shortages_info, key=lambda x: x['shortage_date'])[0]
        model, shortage_date = first_shortage['model'], first_shortage['shortage_date']
        
        proposed_shipping_date = None
        check_date = shortage_date - datetime.timedelta(days=1)
        while self.processor.date_cols and check_date >= self.processor.date_cols[0].date():
            is_shipping_day = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(check_date.weekday()), 'False') == 'True'
            if is_shipping_day and check_date not in self.config_manager.config['NON_SHIPPING_DATES']:
                proposed_shipping_date = check_date
                break
            check_date -= datetime.timedelta(days=1)

        if not proposed_shipping_date:
            messagebox.showwarning("해결 방안 제안 불가", f"{model} 품목의 부족({shortage_date})을 해결할 수 있는 이전 납품일이 없습니다.", parent=self)
            return

        current_max_trucks = self.config_manager.config.get('DAILY_TRUCK_OVERRIDES', {}).get(proposed_shipping_date, self.config_manager.config.get('MAX_TRUCKS_PER_DAY', 2))
        msg = (f"'{model}' 품목의 재고가 {shortage_date.strftime('%m-%d')}에 부족할 것으로 예상됩니다.\n\n"
               f"이 문제를 해결하기 위해, {proposed_shipping_date.strftime('%m-%d')}의 최대 출고 차수를\n"
               f"{current_max_trucks}회에서 {current_max_trucks + 1}회로 늘리시겠습니까?")
        if messagebox.askyesno("재고 부족 해결 방안 제안", msg, parent=self):
            self.config_manager.config['DAILY_TRUCK_OVERRIDES'][proposed_shipping_date] = current_max_trucks + 1
            self.save_settings_and_recalculate()

    def populate_detail_view(self, model_name):
        for widget in self.detail_frame.winfo_children():
            widget.destroy()
        if self.processor.simulated_plan_df is None: return
        model_data = self.processor.simulated_plan_df.loc[model_name]

        fig, ax = plt.subplots(figsize=(12, 4))
        dates = self.processor.date_cols
        date_strs = [d.strftime('%m-%d') for d in dates]
        inventory = [model_data.get(f"재고_{d.strftime('%m%d')}", 0) for d in dates]
        production = [model_data.get(d, 0) for d in dates]
        shipment_cols = [c for c in model_data.index if isinstance(c, str) and c.startswith('출고_')]
        shipments_by_date = {d: 0 for d in dates}
        for col in shipment_cols:
            try:
                date_str_from_col = col[-4:]
                date_obj = datetime.datetime.strptime(f"{dates[0].year}-{date_str_from_col[:2]}-{date_str_from_col[2:]}", "%Y-%m-%d")
                if date_obj in shipments_by_date:
                    shipments_by_date[date_obj] += model_data[col]
            except (ValueError, KeyError): continue
        total_shipments = [shipments_by_date[d] for d in dates]

        ax.plot(date_strs, inventory, marker='o', linestyle='-', label='예상 재고')
        ax.bar(date_strs, production, color='skyblue', label='생산량')
        ax.bar(date_strs, [-s for s in total_shipments], color='salmon', label='총출고량')
        safety_stock = self.processor.item_master_df.loc[model_name, 'SafetyStock']
        if safety_stock > 0:
            ax.axhline(y=safety_stock, color='r', linestyle='--', label=f'최소 재고 ({safety_stock:,})')
        ax.set_title(f"'{model_name}' 재고 및 입출고 추이", fontdict={'fontsize': 14})
        ax.set_xlabel("날짜"); ax.set_ylabel("수량")
        ax.legend(); ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        plt.setp(ax.get_xticklabels(), rotation=45, ha="right"); fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.detail_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True, padx=10, pady=10)
    
    def check_shipment_capacity(self):
        df, messages = self.processor.simulated_plan_df, []
        if df is None or not self.processor.date_cols: return
        truck_capacity = self.config_manager.config.get('PALLETS_PER_TRUCK', 36) * self.config_manager.config.get('PALLET_SIZE', 60)
        all_shipment_cols = [col for col in df.columns if isinstance(col, str) and col.startswith('출고_')]
        grouped_cols = {}
        for col in all_shipment_cols:
            parts = col.split('_'); key = (parts[2], parts[1])
            if key not in grouped_cols: grouped_cols[key] = []
            grouped_cols[key].append(col)
        for (date_str, truck_num), cols in grouped_cols.items():
            total_shipped = df[cols].sum().sum()
            if total_shipped > truck_capacity:
                date_obj = datetime.datetime.strptime(f"{datetime.date.today().year}{date_str}", "%Y%m%d")
                messages.append(f"{date_obj.strftime('%m-%d')} {truck_num}: 출고량 {total_shipped:,.0f} > 용량 {truck_capacity:,.0f}.")
        if messages:
            messagebox.showwarning("출고 용량 초과", "\n".join(messages))
    
    def update_status_bar(self, message="준비 완료"):
        self.status_bar.configure(text=f"현재 파일: {self.current_file} | 상태: {message}")
        logging.info(f"상태 업데이트: {message}")

    def load_settings_to_gui(self):
        for key, entry_widget in self.settings_entries.items():
            entry_widget.delete(0, 'end')
            entry_widget.insert(0, str(self.config_manager.config.get(key, '')))
        for i, cb in self.day_checkboxes.items():
            cb.select() if self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(i), 'False') == 'True' else cb.deselect()
        logging.info("UI에 설정값 로드 완료.")
    
    def open_daily_truck_dialog(self):
        dialog = DailyTruckDialog(self, self.config_manager.config.get('DAILY_TRUCK_OVERRIDES', {}))
        self.wait_window(dialog)
        if dialog.result is not None:
            self.config_manager.config['DAILY_TRUCK_OVERRIDES'] = dialog.result
            self.save_settings_and_recalculate()

    def open_holiday_dialog(self):
        current_holidays = [d for d in self.config_manager.config['NON_SHIPPING_DATES'] if isinstance(d, datetime.date)]
        dialog = HolidayDialog(self, current_holidays)
        self.wait_window(dialog)
        if dialog.result is not None:
            self.config_manager.config['NON_SHIPPING_DATES'] = dialog.result
            self.save_settings_and_recalculate()

    def open_safety_stock_dialog(self):
        if self.processor.item_master_df is None: return
        dialog = SafetyStockDialog(self, self.processor.item_master_df)
        self.wait_window(dialog)
        if dialog.result is not None:
            self.processor.item_master_df = dialog.result
            self.processor.save_item_master()
            messagebox.showinfo("저장 완료", "최소 재고 설정이 저장되었습니다. '설정 저장 및 재계산'으로 계획에 반영하세요.")
            if self.current_step >=2:
                self.recalculate_with_fixed_values()

if __name__ == "__main__":
    try:
        config_manager = ConfigManager()
        app = ProductionPlannerApp(config_manager)
        app.mainloop()
    except Exception as e:
        logging.critical(f"Fatal error: {e}", exc_info=True)
        messagebox.showerror("치명적 오류", f"프로그램 실행에 실패했습니다.\n{e}")