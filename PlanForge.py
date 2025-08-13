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

# ===================================================================
# GitHub 자동 업데이트 설정
# ⚠️ 아래 3개의 변수를 당신의 GitHub 저장소 정보에 맞게 수정하세요!
# ===================================================================
REPO_OWNER = "KMTechn"      # GitHub 사용자 이름 또는 조직 이름
REPO_NAME = "PlanForge"  # 이 프로젝트의 GitHub 저장소 이름
CURRENT_VERSION = "v1.0.0"                   # 현재 프로그램의 버전 (GitHub 릴리스 태그와 일치해야 함)
# ===================================================================


# ===================================================================
# 자동 업데이트 기능 함수들
# ===================================================================
def check_for_updates(repo_owner: str, repo_name: str, current_version: str):
    """
    GitHub에서 최신 릴리스 버전을 확인합니다.
    """
    logging.info("Checking for updates...")
    try:
        api_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
        response = requests.get(api_url, timeout=5)
        response.raise_for_status()
        latest_release_data = response.json()
        latest_version = latest_release_data['tag_name']

        logging.info(f"Current version: {current_version}, Latest version: {latest_version}")

        clean_current = current_version.lower().lstrip('v')
        clean_latest = latest_version.lower().lstrip('v')

        if clean_latest != clean_current:
            for asset in latest_release_data['assets']:
                if asset['name'].endswith('.zip'):
                    return asset['browser_download_url'], latest_version
    except requests.exceptions.RequestException as e:
        logging.error(f"Update check failed: {e}")
    
    return None, None

def download_and_apply_update(url: str):
    """
    업데이트 파일을 다운로드하고, 압축을 해제한 뒤,
    업데이트를 수행하는 배치 파일(updater.bat)을 실행합니다.
    """
    try:
        logging.info(f"Downloading update from: {url}")
        
        temp_dir = os.environ.get("TEMP", "C:\\Temp")
        zip_path = os.path.join(temp_dir, "update.zip")
        
        response = requests.get(url, stream=True, timeout=120)
        response.raise_for_status()
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
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
        root_alert = tk.Tk()
        root_alert.withdraw()
        messagebox.showerror("업데이트 실패", f"업데이트 적용 중 오류가 발생했습니다.\n\n{e}\n\n프로그램을 다시 시작해주세요.", parent=root_alert)
        root_alert.destroy()

def run_updater(repo_owner: str, repo_name: str, current_version: str):
    """
    업데이트를 확인하고 사용자에게 적용 여부를 묻는 메인 함수.
    """
    def check_thread():
        download_url, new_version = check_for_updates(repo_owner, repo_name, current_version)
        if download_url:
            root_alert = tk.Tk()
            root_alert.withdraw()
            if messagebox.askyesno(
                "업데이트 발견", 
                f"새로운 버전({new_version})이 발견되었습니다.\n지금 업데이트하시겠습니까? (현재: {current_version})", 
                parent=root_alert
            ):
                root_alert.destroy()
                download_and_apply_update(download_url)
            else:
                root_alert.destroy()
                logging.info("User declined the update.")
        else:
            logging.info("No new updates found.")

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
        self.item_master_df = None
        self.allowed_models = []
        self.highlight_models = []
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
            df_raw = pd.read_excel(self.current_filepath, sheet_name='《HCO&DIS》', header=None)
            logging.info("원시 데이터 로드 성공. 헤더 행 탐색...")
            part_col_index = 11
            header_row_index = -1
            for i, row in df_raw.iterrows():
                if len(row) > part_col_index and isinstance(row.iloc[part_col_index], str) and 'cover glass assy' in row.iloc[part_col_index].lower():
                    header_row_index = i
                    break
            if header_row_index == -1:
                raise ValueError("헤더 'Cover glass Assy'를 찾을 수 없습니다.")
            
            logging.info(f"헤더 행 발견: {header_row_index}")
            df = df_raw.iloc[header_row_index:].copy()
            df.columns = df.iloc[0]
            df = df.iloc[1:].rename(columns={df.columns[part_col_index]: 'Model'})
            
            self.date_cols = sorted([col for col in df.columns if isinstance(col, (datetime.datetime, pd.Timestamp))])
            
            if not self.date_cols:
                raise ValueError("파일에서 유효한 날짜 컬럼을 찾을 수 없습니다.")

            logging.info(f"유효한 날짜 컬럼 {len(self.date_cols)}개 발견. 모델 필터링 시작...")
            df_filtered = df[df['Model'].isin(self.allowed_models)].copy()
            logging.info(f"유효한 모델로 필터링 후 행 수: {len(df_filtered)}")
            
            df_filtered.loc[:, self.date_cols] = df_filtered.loc[:, self.date_cols].apply(pd.to_numeric, errors='coerce').fillna(0).infer_objects(copy=False)
            agg_df = df_filtered.groupby('Model')[self.date_cols].sum()
            reindexed_df = agg_df.reindex(self.allowed_models).fillna(0).infer_objects(copy=False)
            
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

    def run_simulation(self, adjustments=None, fixed_shipments=None):
        logging.info("시뮬레이션 시작...")
        self.adjustments = adjustments if adjustments else []
        self.fixed_shipments = fixed_shipments if fixed_shipments else []
        if self.aggregated_plan_df is None:
            logging.warning("aggregated_plan_df가 없어 시뮬레이션 중단.")
            return
        
        simulation_dates = self.date_cols[:]
        if self.inventory_date:
            plan_start_date = simulation_dates[0].date() if simulation_dates else None
            if plan_start_date and self.inventory_date > plan_start_date:
                raise ValueError(f"재고 기준일({self.inventory_date.strftime('%Y-%m-%d')})이 생산 계획 시작일({plan_start_date.strftime('%Y-%m-%d')})보다 미래입니다.")
            
            simulation_dates = [d for d in simulation_dates if d.date() >= self.inventory_date]
            if not simulation_dates:
                raise ValueError(f"재고 기준일({self.inventory_date.strftime('%Y-%m-%d')}) 이후에 해당하는 생산 계획이 없습니다.")

        if not simulation_dates:
            logging.warning("시뮬레이션할 유효한 날짜가 없습니다.")
            return

        plan_cols = [col for col in self.aggregated_plan_df.columns if col != 'Status']
        df = self.aggregated_plan_df[plan_cols].copy()
        if self.inventory_df is not None:
            df = df.join(self.inventory_df, how='left').fillna({'Inventory': 0})
        else:
            df = df.assign(Inventory=0)
        df['Inventory'] = df['Inventory'].astype(int)
        
        lead_time = self.config.get('LEAD_TIME_DAYS', 2)
        pallet_size = self.config.get('PALLET_SIZE', 60)
        pallets_per_truck = self.config.get('PALLETS_PER_TRUCK', 36)
        truck_capacity = pallets_per_truck * pallet_size
        
        new_cols = {}
        for date in simulation_dates:
            date_str = date.strftime("%m%d")
            new_cols[f'재고_{date_str}'] = 0
            for t in range(1, 11):
                new_cols[f'출고_{t}차_{date_str}'] = 0
        
        simulated_df = pd.concat([df, pd.DataFrame(columns=list(new_cols.keys()))], axis=1).fillna(0)

        adjustments_by_date = {}
        for adj in self.adjustments:
            date_key = adj['date'].strftime("%Y-%m-%d")
            if date_key not in adjustments_by_date:
                adjustments_by_date[date_key] = []
            adjustments_by_date[date_key].append(adj)
        
        simulated_df['current_inventory'] = simulated_df['Inventory']
        
        for date_idx, date in enumerate(simulation_dates):
            date_str = date.strftime("%m%d")
            date_obj = date.date()
            daily_max_trucks = self.config.get('DAILY_TRUCK_OVERRIDES', {}).get(date_obj, self.config.get('MAX_TRUCKS_PER_DAY', 2))
            
            is_shipping_day = self.config['DELIVERY_DAYS'].get(str(date.weekday()), 'False') == 'True'
            is_non_shipping_date = date.date() in self.config['NON_SHIPPING_DATES']
            
            daily_shipments = {model: {f'출고_{t}차_{date_str}': 0 for t in range(1, daily_max_trucks + 1)} for model in df.index}

            if is_shipping_day and not is_non_shipping_date:
                remaining_capacity_per_truck = [truck_capacity] * daily_max_trucks
            
                fixed_for_day = [s for s in self.fixed_shipments if s['date'] == date.date()]
                for fixed in fixed_for_day:
                    model = fixed['model']
                    shipment_qty = fixed['qty']
                    truck_num = fixed['truck_num']
                    if truck_num <= daily_max_trucks:
                        daily_shipments[model][f'출고_{truck_num}차_{date_str}'] = shipment_qty
                        remaining_capacity_per_truck[truck_num - 1] -= shipment_qty
                
                shortages = []
                for model in df.index:
                    if (date_idx + lead_time) < len(simulation_dates):
                        on_hand_before_prod = simulated_df.loc[model, 'current_inventory']
                        
                        total_production_needed = 0
                        for i in range(lead_time + 1):
                            if (date_idx + i) < len(simulation_dates):
                                total_production_needed += df.loc[model, simulation_dates[date_idx+i]]
                        
                        safety_stock = self.item_master_df.loc[model, 'SafetyStock']
                        total_shipped_fixed = sum(daily_shipments[model].values())
                        
                        required = max(0, (total_production_needed + safety_stock) - (on_hand_before_prod + total_shipped_fixed))
                        
                        if required > 0:
                            shortages.append({'model': model, 'required': required})
                shortages.sort(key=lambda x: x['required'], reverse=True)
                
                for shortage in shortages:
                    model = shortage['model']
                    remaining_to_ship = shortage['required']
                    for truck_num in range(daily_max_trucks):
                        if remaining_capacity_per_truck[truck_num] > 0 and remaining_to_ship > 0:
                            shipment_needed = remaining_to_ship
                            shipment = math.ceil(shipment_needed / pallet_size) * pallet_size
                            shipment = min(shipment, remaining_capacity_per_truck[truck_num])
                            if shipment > 0:
                                daily_shipments[model][f'출고_{truck_num+1}차_{date_str}'] += shipment
                                remaining_capacity_per_truck[truck_num] -= shipment
                                remaining_to_ship = max(0, remaining_to_ship - shipment)
                
                if any(cap > 0 for cap in remaining_capacity_per_truck):
                    sorted_models = self.item_master_df.sort_values('Priority').index.tolist()
                    for model in sorted_models:
                        if sum(daily_shipments[model].values()) > 0: continue
                        
                        on_hand_before_prod = simulated_df.loc[model, 'current_inventory']
                        
                        future_demand = 0
                        for i in range(lead_time + 1, len(simulation_dates) - date_idx):
                            future_demand += df.loc[model, simulation_dates[date_idx+i]]
                        
                        if future_demand > 0:
                            total_shipped_fixed = sum(daily_shipments[model].values())
                            required = max(0, future_demand - (on_hand_before_prod + total_shipped_fixed))
                            
                            if required > 0:
                                remaining_proactive = required
                                for truck_num in range(daily_max_trucks):
                                    if remaining_capacity_per_truck[truck_num] > 0:
                                        proactive_needed = remaining_proactive
                                        proactive_shipment = math.ceil(proactive_needed / pallet_size) * pallet_size
                                        proactive_shipment = min(proactive_shipment, remaining_capacity_per_truck[truck_num])
                                        if proactive_shipment > 0:
                                            daily_shipments[model][f'출고_{truck_num+1}차_{date_str}'] += proactive_shipment
                                            remaining_capacity_per_truck[truck_num] -= proactive_shipment
                                            remaining_proactive = max(0, remaining_proactive - proactive_shipment)
                                            if remaining_proactive <= 0:
                                                break
                                if all(cap == 0 for cap in remaining_capacity_per_truck):
                                    break

            for model in df.index:
                today_production = df.loc[model, date] if date in df.columns else 0
                inventory_adjustment = 0
                for adj in adjustments_by_date.get(date.date().strftime("%Y-%m-%d"), []):
                    if adj['model'] == model:
                        if adj['type'] == '수요':
                            today_production += adj['qty']
                        elif adj['type'] == '재고':
                            inventory_adjustment += adj['qty']
                total_shipped = sum(daily_shipments[model].values())
                on_hand = simulated_df.loc[model, 'current_inventory']
                new_inventory = on_hand + total_shipped + inventory_adjustment - today_production
                
                simulated_df.loc[model, f'재고_{date_str}'] = new_inventory
                for k, v in daily_shipments[model].items():
                    if k in simulated_df.columns:
                        simulated_df.loc[model, k] = v
                
                simulated_df.loc[model, 'current_inventory'] = new_inventory
        
        self.simulated_plan_df = simulated_df.drop(columns=['current_inventory']).astype(float).fillna(0).astype(int)
        logging.info("시뮬레이션 완료.")

class AdjustmentDialog(ctk.CTkToplevel):
    def __init__(self, parent, models):
        super().__init__(parent)
        self.models = models
        self.adjustments = []
        self.result = None
        self.title("수동 조정 입력")
        self.geometry("600x450")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
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
        self.listbox.insert(END, f"{adj['date']}, {adj['model']}, {adj['qty']}, {adj['type']}")
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
        
        self.sidebar_visible = True

        self.create_widgets()
        self.update_status_bar()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind_all("<Control-MouseWheel>", self.on_mouse_wheel_zoom)
        self.inventory_text_backup = None
        self.after_ids = []
        self.last_selected_model = None
        
        # 앱 시작 시 업데이트 확인
        run_updater(REPO_OWNER, REPO_NAME, CURRENT_VERSION)

    def on_closing(self):
        try:
            for after_id in self.after_ids:
                self.after_cancel(after_id)
            self.after_ids = []
            self.unbind_all("<Control-MouseWheel>")
            plt.close('all')
            if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
                self.destroy()
        except Exception as e:
            logging.error(f"Closing error: {e}")
            self.destroy()

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
        main_area_frame.grid_rowconfigure(2, weight=1)

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
        ctk.CTkLabel(self.delivery_days_frame, text="납품 요일:", font=self.font_normal).pack(anchor="w")
        self.day_checkboxes = {}
        day_names = ["월", "화", "수", "목", "금", "토", "일"]
        for i, day in enumerate(day_names):
            state = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(i), 'False') == 'True'
            cb = ctk.CTkCheckBox(self.delivery_days_frame, text=day, onvalue=True, offvalue=False, font=self.font_normal)
            cb.pack(side='left', padx=2)
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
        
        self.shortage_frame = ctk.CTkFrame(main_area_frame, fg_color="#FFF5E1")
        self.shortage_frame.grid(row=1, column=0, sticky="ew", pady=5)
        self.shortage_frame.grid_remove()
        shortage_title = ctk.CTkLabel(self.shortage_frame, text="⚠️ 재고 부족 및 해결 방안 제시", font=self.font_bold, text_color="#E67E22")
        shortage_title.pack(pady=(5,0))
        self.shortage_list_frame = ctk.CTkScrollableFrame(self.shortage_frame, label_text="", height=100)
        self.shortage_list_frame.pack(fill="x", expand=True, padx=5, pady=5)

        self.tabview = ctk.CTkTabview(main_area_frame)
        self.tabview.grid(row=2, column=0, sticky="nsew")
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

    def update_shortage_warnings(self):
        for widget in self.shortage_list_frame.winfo_children():
            widget.destroy()

        df = self.processor.simulated_plan_df
        if df is None:
            self.shortage_frame.grid_remove()
            return

        inventory_cols = sorted([col for col in df.columns if isinstance(col, str) and col.startswith('재고_')])
        shortages = []
        for col in inventory_cols:
            date_str = col.replace('재고_', '')
            date_obj = datetime.datetime.strptime(f"{datetime.date.today().year}{date_str}", "%Y%m%d").date()
            
            for model, qty in df[col].items():
                safety_stock = self.processor.item_master_df.loc[model, 'SafetyStock']
                if qty < safety_stock:
                    shortages.append({'date': date_obj, 'model': model, 'qty': qty, 'shortage_qty': safety_stock - qty})
        
        if shortages:
            self.shortage_frame.grid()
            
            for s in sorted(shortages, key=lambda x: x['date']):
                msg_frame = ctk.CTkFrame(self.shortage_list_frame, fg_color="transparent")
                msg_frame.pack(fill="x", padx=5, pady=3)
                
                msg = f"[{s['date'].strftime('%m-%d')}] '{s['model']}' 재고 {s['qty']:,}개. (안전재고까지 {s['shortage_qty']:,}개 부족)"
                lbl = ctk.CTkLabel(msg_frame, text=msg, font=self.font_bold, text_color="#D35400", anchor="w")
                lbl.pack(fill="x")
                lbl.bind("<Double-Button-1>", lambda e, m=s['model']: self.on_row_double_click(m))

        else:
            self.shortage_frame.grid_remove()

    def prompt_for_font_size(self, event=None):
        dialog = ctk.CTkInputDialog(text="새로운 폰트 크기를 입력하세요 (5-50):", title="폰트 크기 변경")
        new_size_str = dialog.get_input()
        if new_size_str:
            try:
                new_size = int(new_size_str)
                if not (5 <= new_size <= 50):
                    self.update_status_bar("폰트 크기는 5와 50 사이의 숫자여야 합니다.")
                else:
                    self.set_font_size(new_size)
            except (ValueError, TypeError):
                self.update_status_bar("유효한 숫자를 입력해주세요.")

    def set_font_size(self, new_size):
        if new_size == self.base_font_size:
            return
        self.base_font_size = new_size
        self.font_normal.configure(size=self.base_font_size)
        self.font_small.configure(size=self.base_font_size - 1)
        self.font_bold.configure(size=self.base_font_size, weight="bold")
        self.font_italic.configure(size=self.base_font_size, slant="italic")
        self.font_header.configure(size=self.base_font_size + 1, weight="bold")
        self.font_edit.configure(size=self.base_font_size, weight="bold")
        self.update_static_fonts()

    def change_font_size(self, delta):
        new_size = max(5, min(50, self.base_font_size + delta))
        self.set_font_size(new_size)

    def update_static_fonts(self):
        self.sidebar_title.configure(font=self.font_big_bold)
        self.step1_button.configure(font=self.font_normal)
        self.step2_button.configure(font=self.font_normal)
        self.step3_button.configure(font=self.font_normal)
        self.step4_button.configure(font=self.font_normal)
        self.font_size_title_label.configure(font=self.font_normal)
        self.font_minus_button.configure(font=self.font_normal)
        self.font_size_label.configure(font=self.font_normal, text=str(self.base_font_size))
        self.font_plus_button.configure(font=self.font_normal)
        self.settings_title_label.configure(font=self.font_bold)
        for label in self.setting_labels:
            label.configure(font=self.font_normal)
        for entry in self.settings_entries.values():
            entry.configure(font=self.font_normal)
        self.save_settings_button.configure(font=self.font_normal)
        self.search_label.configure(font=self.font_normal)
        self.search_entry.configure(font=self.font_normal)
        self.lbl_models_found.configure(font=self.font_kpi)
        self.lbl_total_quantity.configure(font=self.font_kpi)
        self.lbl_date_range.configure(font=self.font_kpi)
        self.detail_tab_title.configure(font=self.font_bold)
        self.tabview.configure(font=self.font_normal)
        self.filter_grid()
        if self.current_step >= 2 and hasattr(self, 'last_selected_model') and self.last_selected_model:
            self.populate_detail_view(self.last_selected_model)

    def populate_master_grid(self, df_to_show):
        for widget in self.master_frame.winfo_children():
            widget.destroy()
        if df_to_show is None or df_to_show.empty:
            return
        
        df_to_show = df_to_show.infer_objects(copy=False)
        self.master_frame.grid_columnconfigure(0, minsize=140)

        plan_cols = self.processor.date_cols
        
        if 'Model' not in df_to_show.columns:
            df_display = df_to_show.reset_index().rename(columns={'index': 'Model'})
        else:
            df_display = df_to_show

        if self.current_step < 2:
            headers = ['Model'] + [d.strftime('%m-%d') for d in plan_cols]
            for c, h_text in enumerate(headers):
                self.master_frame.grid_columnconfigure(c, weight=1 if c > 0 else 0)
                ctk.CTkLabel(self.master_frame, text=h_text, font=self.font_header, anchor="center").grid(row=0, column=c, sticky="ew", padx=1, pady=2)
            
            for r, row_data in df_display.iterrows():
                model = row_data['Model']
                is_highlighted = model in self.processor.highlight_models
                bg_color = "#D6EAF8" if is_highlighted else "transparent"
                
                lbl_model = ctk.CTkLabel(self.master_frame, text=model, fg_color=bg_color, font=self.font_normal, anchor="w", padx=5)
                lbl_model.grid(row=r + 1, column=0, sticky="ew")
                lbl_model.bind("<Double-Button-1>", lambda e, m=model: self.on_row_double_click(m))

                for i, date_col in enumerate(plan_cols):
                    val = row_data.get(date_col, 0)
                    text = f"{val:,.0f}" if val else "0"
                    lbl_data = ctk.CTkLabel(self.master_frame, text=text, fg_color=bg_color, font=self.font_normal, anchor="e", padx=5)
                    lbl_data.grid(row=r + 1, column=i + 1, sticky="ew")
                    lbl_data.bind("<Double-Button-1>", lambda e, m=model: self.on_row_double_click(m))
            
            totals = df_to_show[plan_cols].sum()
            ctk.CTkFrame(self.master_frame, height=1, fg_color="lightgray").grid(row=len(df_display)+1, column=0, columnspan=len(headers), sticky='ew', pady=4)
            ctk.CTkLabel(self.master_frame, text="합계", font=self.font_bold, anchor="w", padx=5).grid(row=len(df_display)+2, column=0, sticky="ew")
            for i, date_col in enumerate(plan_cols):
                total_val = totals.get(date_col, 0)
                ctk.CTkLabel(self.master_frame, text=f"{total_val:,.0f}", font=self.font_bold, anchor="e", padx=5).grid(row=len(df_display)+2, column=i+1, sticky="ew")

        else: # Step 2 이후
            max_trucks_default = self.config_manager.config.get('MAX_TRUCKS_PER_DAY', 2)
            
            ctk.CTkLabel(self.master_frame, text="Model", font=self.font_header, anchor="center").grid(row=0, column=0, rowspan=2, sticky="nsew", padx=1, pady=2)
            
            current_col_idx = 1
            col_idx_map = {}

            for d in plan_cols:
                daily_max = self.config_manager.config.get('DAILY_TRUCK_OVERRIDES', {}).get(d.date(), max_trucks_default)
                ship_cols = [c for c in df_to_show.columns if isinstance(c, str) and d.strftime("%m%d") in c and c.startswith('출고_')]
                used_trucks = len([c for c in ship_cols if df_to_show[c].sum() > 0])
                
                date_header_text = f"{d.strftime('%m-%d')}\n({used_trucks}/{daily_max}대)"
                ctk.CTkLabel(self.master_frame, text=date_header_text, font=self.font_header, anchor="center", justify="center").grid(row=0, column=current_col_idx, columnspan=daily_max, sticky="ew", padx=1, pady=2)
                
                col_idx_map[d.date()] = current_col_idx

                for truck_num in range(1, daily_max + 1):
                    sub_header_text = f"{truck_num}차"
                    ctk.CTkLabel(self.master_frame, text=sub_header_text, font=self.font_header, anchor="center").grid(row=1, column=current_col_idx, sticky="ew", padx=1, pady=2)
                    current_col_idx += 1
            
            for r, row_data in df_display.iterrows():
                row_idx = r + 2
                model = row_data['Model']
                is_highlighted = model in self.processor.highlight_models
                bg_color = "#D6EAF8" if is_highlighted else "#FFFFFF"
                
                lbl_model = ctk.CTkLabel(self.master_frame, text=model, fg_color=bg_color, font=self.font_normal, anchor="w", padx=5)
                lbl_model.grid(row=row_idx, column=0, sticky="ew")
                lbl_model.bind("<Double-Button-1>", lambda e, m=model: self.on_row_double_click(m))

                for date_col in plan_cols:
                    start_col = col_idx_map[date_col.date()]
                    daily_max = self.config_manager.config.get('DAILY_TRUCK_OVERRIDES', {}).get(date_col.date(), max_trucks_default)
                    is_shipping_day = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(date_col.weekday()), 'False') == 'True'
                    is_non_shipping_date = date_col.date() in self.config_manager.config['NON_SHIPPING_DATES']
                    is_non_shipping_day_or_date = not is_shipping_day or is_non_shipping_date

                    for truck_num in range(1, daily_max + 1):
                        col_name = f'출고_{truck_num}차_{date_col.strftime("%m%d")}'
                        val = row_data.get(col_name, 0)
                        is_fixed = any(s['model'] == model and s['date'] == date_col.date() and s['truck_num'] == truck_num for s in self.processor.fixed_shipments)
                        
                        text = f"{val:,.0f}" if val else "0"
                        label_bg_color = bg_color
                        if is_fixed: label_bg_color = "#A9CCE3"
                        if is_non_shipping_day_or_date:
                            label_bg_color = "#F2F3F4"
                            text = "-"

                        data_label = ctk.CTkLabel(self.master_frame, text=text, fg_color=label_bg_color, font=self.font_bold if is_fixed else self.font_normal, anchor="e", padx=5, text_color="blue" if is_fixed else "black")
                        data_label.grid(row=row_idx, column=start_col + truck_num - 1, sticky="ew")
                        
                        if not is_non_shipping_day_or_date:
                            data_label.bind("<Double-Button-1>", lambda e, m=model, d=date_col.date(), t=truck_num: self.on_shipment_double_click(e, m, d, t))
                            data_label.bind("<Button-3>", lambda e, m=model, d=date_col.date(), t=truck_num: self.on_shipment_right_click(e, m, d, t))
            
            total_row_idx = len(df_display) + 2
            total_cols = current_col_idx 
            ctk.CTkFrame(self.master_frame, height=1, fg_color="lightgray").grid(row=total_row_idx, column=0, columnspan=total_cols, sticky='ew', pady=4)
            ctk.CTkLabel(self.master_frame, text="합계", font=self.font_bold, anchor="w", padx=5).grid(row=total_row_idx + 1, column=0, sticky="ew")

            for date_col in plan_cols:
                start_col = col_idx_map[date_col.date()]
                daily_max = self.config_manager.config.get('DAILY_TRUCK_OVERRIDES', {}).get(date_col.date(), max_trucks_default)
                for truck_num in range(1, daily_max + 1):
                    col_name = f'출고_{truck_num}차_{date_col.strftime("%m%d")}'
                    total_val = df_display[col_name].sum() if col_name in df_display else 0
                    ctk.CTkLabel(self.master_frame, text=f"{total_val:,.0f}", font=self.font_bold, anchor="e", padx=5).grid(row=total_row_idx + 1, column=start_col + truck_num - 1, sticky="ew")

    def populate_detail_view(self, model_name):
        for widget in self.detail_frame.winfo_children():
            widget.destroy()
        df = self.processor.simulated_plan_df
        if df is None or model_name not in df.index:
            ctk.CTkLabel(self.detail_frame, text=f"'{model_name}'에 대한 시뮬레이션 데이터를 찾을 수 없습니다.", font=self.font_normal).pack(pady=20)
            return
        row_data = df.loc[model_name]
        self.detail_tab_title.configure(text=f"상세: '{model_name}' 출고 시뮬레이션")
        date_cols = self.processor.date_cols
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
            if 'weight' in font_opts and font_opts['weight'] == 'bold':
                label_font = self.font_bold
            data_font = self.font_bold if 'weight' in font_opts else self.font_normal
            ctk.CTkLabel(self.detail_frame, text=name, anchor="w", font=label_font, **label_opts).grid(row=current_row, column=0, sticky="w", padx=5)
            if is_initial:
                val = row_data.get(key_prefix, 0)
                ctk.CTkLabel(self.detail_frame, text=f"{val:,.0f}", font=data_font, **label_opts).grid(row=current_row, column=1, sticky="w")
            else:
                for c, date_col in enumerate(date_cols):
                    val = row_data.get(date_col, 0) if is_date_key else row_data.get(f'{key_prefix}{date_col.strftime("%m%d")}', 0)
                    
                    day_specific_opts = label_opts.copy()
                    
                    if key_prefix == '재고_':
                        safety_stock = self.processor.item_master_df.loc[model_name, 'SafetyStock']
                        if val < 0:
                            day_specific_opts['text_color'] = 'red'
                        elif val < safety_stock:
                            day_specific_opts['text_color'] = 'orange'


                    ctk.CTkLabel(self.detail_frame, text=f"{val:,.0f}", font=data_font, **day_specific_opts, anchor="e").grid(row=current_row, column=c+1, padx=5, sticky="ew")
            current_row += 1

        def draw_separator():
            nonlocal current_row
            ctk.CTkFrame(self.detail_frame, height=1, fg_color="lightgray").grid(row=current_row, column=0, columnspan=len(headers), sticky="ew", pady=4)
            current_row += 1
        
        draw_row("초기 재고", "Inventory", options={'weight':'bold'}, is_initial=True)
        draw_row("출고 (생산)", "", options={'text_color':"red"}, is_date_key=True)
        draw_separator()
        
        all_truck_nums = set()
        for col in df.columns:
            if isinstance(col, str) and col.startswith("출고_"):
                try:
                    num = int(col.split("_")[1].replace("차",""))
                    all_truck_nums.add(num)
                except (ValueError, IndexError):
                    continue
        
        for i in sorted(list(all_truck_nums)):
            if row_data[[c for c in df.columns if isinstance(c, str) and f'출고_{i}차' in c]].sum() > 0:
                draw_row(f"{i}차 출고", f"출고_{i}차_", options={'weight':'bold', 'text_color':"#2E86C1"})
        
        draw_separator()
        draw_row("일일 재고", "재고_", options={'weight':'bold'})
        
        fig, ax = plt.subplots(figsize=(8, 3))
        inventory_vals = [row_data.get(f'재고_{d.strftime("%m%d")}', 0) for d in date_cols]
        labels = [d.strftime('%m-%d') for d in date_cols]
        ax.plot(labels, inventory_vals, marker='o', color="#3498DB", label="재고량")
        
        safety_stock_val = self.processor.item_master_df.loc[model_name, 'SafetyStock']
        ax.axhline(safety_stock_val, color='orange', linestyle='--', linewidth=1.5, label=f"안전재고 ({safety_stock_val:,})")

        ax.set_title(f"'{model_name}' 일일 재고 추이", fontdict={'fontsize': 10})
        ax.set_xlabel('날짜', fontdict={'fontsize': 9})
        ax.set_ylabel('재고량', fontdict={'fontsize': 9})
        ax.axhline(0, color='red', linestyle='--', linewidth=1)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend()
        plt.setp(ax.get_xticklabels(), rotation=45, ha="right", rotation_mode="anchor", fontsize=8)
        plt.setp(ax.get_yticklabels(), fontsize=8)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.detail_frame)
        canvas.draw()
        canvas.get_tk_widget().grid(row=current_row, column=0, columnspan=len(headers), pady=10)

    def run_step1_aggregate(self):
        logging.info("1단계: 생산계획 불러오기 시작.")
        if not os.path.exists(self.config_manager.config_path):
            self.config_manager.save_config(self.config_manager.config)
            self.update_status_bar("새로운 설정 파일 'config.xlsx'가 생성되었습니다.")
        else:
            self.config_manager.load_config()
            self.load_settings_to_gui()

        file_path = filedialog.askopenfilename(title="생산계획 엑셀 파일 선택", filetypes=(("Excel", "*.xlsx *.xls"),))
        if not file_path:
            logging.info("사용자가 파일 선택을 취소했습니다.")
            return

        try:
            self.processor.current_filepath = file_path
            self.processor.process_plan_file()
            self.current_file = os.path.basename(file_path)
            self.current_step = 1

            if self.processor.aggregated_plan_df is None or self.processor.aggregated_plan_df.empty:
                messagebox.showinfo("정보", "처리할 생산 계획 데이터가 없습니다.")
                logging.warning("집계된 데이터가 비어 있습니다.")
                return

            plan_cols = self.processor.date_cols
            df = self.processor.aggregated_plan_df
            df_filtered = df[df[plan_cols].sum(axis=1) > 0]
            models_found = len(df_filtered.index)
            total_qty = df_filtered[plan_cols].sum().sum()

            date_range = f"{plan_cols[0].strftime('%y/%m/%d')} ~ {plan_cols[-1].strftime('%y/%m/%d')}"

            self.lbl_models_found.configure(text=f"처리된 모델 수: {models_found} 개")
            self.lbl_total_quantity.configure(text=f"총생산량: {total_qty:,.0f} 개")
            self.lbl_date_range.configure(text=f"계획 기간: {date_range}")
            self.filter_grid()
            [widget.destroy() for widget in self.detail_frame.winfo_children()]
            self.update_status_bar("1단계: 생산계획 집계 완료")
            self.step2_button.configure(state="normal")
            self.shortage_frame.grid_remove()
            logging.info("1단계 완료. UI 업데이트 완료.")
        except Exception as e:
            messagebox.showerror("1단계 파일 처리 실패", f"{e}")
            logging.error(f"1단계 실행 중 오류 발생: {e}", exc_info=True)

    def run_step2_simulation(self):
        logging.info("2단계: 시뮬레이션 시작.")
        
        dialog = InventoryInputDialog(self)
        self.wait_window(dialog)
        result = dialog.result

        if not result:
            logging.info("사용자가 재고 데이터 입력을 취소했습니다.")
            return

        source_type, param = result
        
        try:
            if source_type == 'text':
                self.processor.load_inventory_from_text(param)
                self.inventory_text_backup = param
            elif source_type == 'file':
                self.processor.load_inventory_from_file(param)
                self.inventory_text_backup = None
            
            self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments)
            
            if self.processor.simulated_plan_df is None:
                logging.warning("시뮬레이션 결과가 생성되지 않았습니다.")
                messagebox.showwarning("시뮬레이션 오류", "시뮬레이션 결과가 생성되지 않았습니다. 입력값을 확인해주세요.")
                return

            self.current_step = 2
            total_ship = self.processor.simulated_plan_df[[col for col in self.processor.simulated_plan_df.columns if isinstance(col, str) and col.startswith('출고_')]].sum().sum()
            self.lbl_total_quantity.configure(text=f"총출고량: {total_ship:,.0f} 개")
            self.filter_grid()
            [widget.destroy() for widget in self.detail_frame.winfo_children()]
            self.update_status_bar("2단계: 출고 계획 시뮬레이션 완료.")
            self.step3_button.configure(state="normal")
            self.step4_button.configure(state="normal")
            self.check_shipment_capacity()
            self.update_shortage_warnings()
            logging.info("2단계 완료. 시뮬레이션 결과 UI 업데이트 완료.")
        except Exception as e:
            messagebox.showerror("2단계 시뮬레이션 실패", f"{e}")
            logging.error(f"2단계 실행 중 오류 발생: {e}", exc_info=True)

    def run_step3_adjustments(self):
        logging.info("3단계: 수동 조정 시작.")
        if self.current_step < 2:
            messagebox.showwarning("오류", "먼저 2단계(재고 반영)를 실행해야 합니다.")
            return
        dialog = AdjustmentDialog(self, models=self.processor.allowed_models)
        self.wait_window(dialog)
        adjustments = dialog.result
        if adjustments is None:
            logging.info("사용자가 수동 조정 입력을 취소했습니다.")
            return
        try:
            if self.inventory_text_backup:
                self.processor.load_inventory_from_text(self.inventory_text_backup)
            
            self.processor.run_simulation(adjustments=adjustments, fixed_shipments=self.processor.fixed_shipments)
            self.current_step = 3
            total_ship = self.processor.simulated_plan_df[[col for col in self.processor.simulated_plan_df.columns if isinstance(col, str) and col.startswith('출고_')]].sum().sum()
            self.lbl_total_quantity.configure(text=f"총출고량: {total_ship:,.0f} 개")
            self.filter_grid()
            [widget.destroy() for widget in self.detail_frame.winfo_children()]
            self.update_status_bar("3단계: 수동 조정 적용 완료.")
            self.check_shipment_capacity()
            self.update_shortage_warnings()
            logging.info("3단계 완료. 조정 결과 UI 업데이트 완료.")
        except Exception as e:
            messagebox.showerror("3단계 조정 실패", f"{e}")
            logging.error(f"3단계 실행 중 오류 발생: {e}", exc_info=True)
    
    def check_shipment_capacity(self):
        df = self.processor.simulated_plan_df
        if df is None or not self.processor.date_cols:
            return
        
        truck_capacity = self.config_manager.config.get('PALLETS_PER_TRUCK', 36) * self.config_manager.config.get('PALLET_SIZE', 60)
        messages = []
        
        all_shipment_cols = [col for col in df.columns if isinstance(col, str) and col.startswith('출고_')]
        
        grouped_cols = {}
        for col in all_shipment_cols:
            parts = col.split('_')
            truck_num = parts[1]
            date_str = parts[2]
            key = (date_str, truck_num)
            if key not in grouped_cols:
                grouped_cols[key] = []
            grouped_cols[key].append(col)

        for (date_str, truck_num), cols in grouped_cols.items():
            total_shipped = df[cols].sum().sum()
            if total_shipped > truck_capacity:
                date_obj = datetime.datetime.strptime(f"{datetime.date.today().year}{date_str}", "%Y%m%d")
                messages.append(f"{date_obj.strftime('%m-%d')} {truck_num}: 출고량 {total_shipped:,.0f} > 용량 {truck_capacity:,.0f}.")
        
        if messages:
            messagebox.showwarning("출고 용량 초과", "\n".join(messages))
            logging.warning("출고 용량 초과 경고 발생.")

    def on_row_double_click(self, model_name):
        logging.info(f"모델 행 더블 클릭됨: {model_name}")
        
        if self.current_step < 2:
            self.update_status_bar("상세 뷰를 보려면 2단계 시뮬레이션을 먼저 실행해야 합니다.")
            self.bell()
            return

        self.last_selected_model = model_name
        
        self.populate_detail_view(model_name)
        
        self.tabview.set("상세")
        
        self.update_status_bar(f"'{model_name}'의 상세 정보를 표시합니다.")

    def on_shipment_double_click(self, event, model, date, truck_num):
        if self.current_step < 2: return
        
        is_shipping_day = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(date.weekday()), 'False') == 'True'
        is_non_shipping_date = date in self.config_manager.config['NON_SHIPPING_DATES']
        if not is_shipping_day or is_non_shipping_date:
            messagebox.showinfo("출고 불가", "납품 불가능한 요일이거나 휴무일입니다.")
            return
        
        is_fixed = any(s['model'] == model and s['date'] == date and s['truck_num'] == truck_num for s in self.processor.fixed_shipments)
        if is_fixed:
            messagebox.showinfo("수정 불가", "고정된 항목은 더블클릭으로 수정할 수 없습니다. 우클릭 메뉴를 이용해주세요.")
            return
        
        pallet_size = self.config_manager.config.get('PALLET_SIZE', 60)
        
        dialog = ctk.CTkInputDialog(text=f"'{model}'의 {date.strftime('%m-%d')} {truck_num}차 출고량을 수정하세요.\n(팔레트 단위 {pallet_size}의 배수로 입력 권장)", title="출고량 수정")
        new_value_str = dialog.get_input()
        
        if new_value_str:
            try:
                new_value = int(new_value_str)
                if new_value < 0: raise ValueError
                
                self.fix_shipment(model, date, new_value, truck_num)
                self.recalculate_with_fixed_values()
                self.update_status_bar(f"'{model}'의 {date.strftime('%m-%d')} {truck_num}차 출고량이 {new_value:,.0f}개로 수동 조정되었습니다.")
                logging.info(f"출고량 수동 조정: 모델='{model}', 날짜={date}, 차수={truck_num}, 수량={new_value}")
            except ValueError:
                messagebox.showerror("입력 오류", "유효한 양의 숫자를 입력해주세요.")

    def on_shipment_right_click(self, event, model, date, truck_num):
        if self.current_step < 2: return
        
        is_shipping_day = self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(date.weekday()), 'False') == 'True'
        is_non_shipping_date = date in self.config_manager.config['NON_SHIPPING_DATES']
        if not is_shipping_day or is_non_shipping_date: return
        
        menu = Menu(self, tearoff=0)
        is_fixed = any(s['model'] == model and s['date'] == date and s['truck_num'] == truck_num for s in self.processor.fixed_shipments)
        
        if is_fixed:
            menu.add_command(label=f"{truck_num}차 고정 해제", command=lambda: self.unfix_shipment(model, date, truck_num))
        else:
            menu.add_command(label=f"{truck_num}차 고정", command=lambda: self.on_fix_request(model, date, truck_num))
        
        menu.tk_popup(event.x_root, event.y_root)

    def on_fix_request(self, model, date, truck_num):
        if self.processor.simulated_plan_df is None: return
        
        col_name = f'출고_{truck_num}차_{date.strftime("%m%d")}'
        shipment_value = self.processor.simulated_plan_df.loc[model, col_name]
        
        self.fix_shipment(model, date, shipment_value, truck_num)
        self.recalculate_with_fixed_values()
        self.update_status_bar(f"'{model}'의 {date.strftime('%m-%d')} {truck_num}차 출고량이 {shipment_value:,.0f}개로 고정되었습니다.")
        logging.info(f"출고량 고정: 모델='{model}', 날짜={date}, 차수={truck_num}, 수량={shipment_value}")

    def fix_shipment(self, model, date, qty, truck_num):
        self.processor.fixed_shipments = [s for s in self.processor.fixed_shipments if not (s['model'] == model and s['date'] == date and s['truck_num'] == truck_num)]
        self.processor.fixed_shipments.append({'model': model, 'date': date, 'qty': qty, 'truck_num': truck_num})
        logging.info(f"출고량 고정 업데이트: 모델='{model}', 날짜={date}, 수량={qty}, 차수={truck_num}")

    def unfix_shipment(self, model, date, truck_num):
        self.processor.fixed_shipments = [s for s in self.processor.fixed_shipments if not (s['model'] == model and s['date'] == date and s['truck_num'] == truck_num)]
        self.recalculate_with_fixed_values()
        self.update_status_bar(f"'{model}'의 {date.strftime('%m-%d')} {truck_num}차 출고량이 고정 해제되었습니다.")
        logging.info(f"출고량 고정 해제: 모델='{model}', 날짜={date}, 차수={truck_num}")

    def recalculate_with_fixed_values(self):
        logging.info("고정값 적용 후 재계산 시작...")
        try:
            if self.inventory_text_backup:
                self.processor.load_inventory_from_text(self.inventory_text_backup)

            self.processor.run_simulation(adjustments=self.processor.adjustments, fixed_shipments=self.processor.fixed_shipments)
            
            if self.processor.simulated_plan_df is None:
                messagebox.showerror("재계산 실패", "시뮬레이션 결과 생성에 실패했습니다.")
                return

            total_ship = self.processor.simulated_plan_df[[col for col in self.processor.simulated_plan_df.columns if isinstance(col, str) and col.startswith('출고_')]].sum().sum()
            self.lbl_total_quantity.configure(text=f"총출고량: {total_ship:,.0f} 개")
            self.filter_grid()
            if hasattr(self, 'last_selected_model') and self.last_selected_model:
                self.populate_detail_view(self.last_selected_model)
            self.update_shortage_warnings()
            logging.info("재계산 완료. UI 업데이트 완료.")
        except Exception as e:
            messagebox.showerror("재계산 실패", f"재계산 중 오류 발생: {e}")
            logging.error(f"재계산 중 오류 발생: {e}", exc_info=True)
    
    def export_to_excel(self):
        if self.current_step == 0:
            messagebox.showwarning("오류", "먼저 1단계 '생산계획 불러오기'를 진행해야 합니다.")
            return
            
        start_date = self.processor.date_cols[0].strftime('%m-%d')
        end_date = self.processor.date_cols[-1].strftime('%m-%d')
        if self.current_step == 1:
            df_to_export = self.processor.aggregated_plan_df
            filename = f"{start_date}~{end_date} 생산계획.xlsx"
        elif self.current_step >= 2:
            df_to_export = self.processor.simulated_plan_df
            filename = f"{start_date}~{end_date} 출고계획.xlsx"
        else:
            messagebox.showwarning("오류", "내보낼 데이터가 없습니다.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=filename, filetypes=(("Excel", "*.xlsx"),))
        if not file_path:
            return
            
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                if self.current_step == 1:
                    df_to_export_filtered = df_to_export[df_to_export[self.processor.date_cols].sum(axis=1) > 0]
                    df_to_export_filtered.to_excel(writer, sheet_name='생산계획')
                else:
                    shipment_cols = [col for col in df_to_export.columns if isinstance(col, str) and col.startswith('출고_')]
                    df_to_export_filtered = df_to_export[df_to_export[shipment_cols].sum(axis=1) > 0]
                    df_to_export_filtered.to_excel(writer, sheet_name='Full Plan')

                    all_truck_nums = set()
                    for col in df_to_export_filtered.columns:
                        if isinstance(col, str) and col.startswith("출고_"):
                            try:
                                num = int(col.split("_")[1].replace("차",""))
                                all_truck_nums.add(num)
                            except (ValueError, IndexError):
                                continue

                    for truck_num in sorted(list(all_truck_nums)):
                        sheet_name_truck = f'{truck_num}차 출고'
                        df_round = pd.DataFrame(index=df_to_export_filtered.index, columns=self.processor.date_cols)
                        for date in self.processor.date_cols:
                            date_str = date.strftime("%m%d")
                            col = f'출고_{truck_num}차_{date_str}'
                            if col in df_to_export_filtered.columns:
                                df_round[date] = df_to_export_filtered[col]
                        df_round.columns = [d.strftime('%m-%d') for d in self.processor.date_cols]
                        df_round.to_excel(writer, sheet_name=sheet_name_truck)
            messagebox.showinfo("내보내기 성공", f"계획이 {file_path}로 저장되었습니다.")
            logging.info(f"계획을 {file_path}로 성공적으로 내보냈습니다.")
        except Exception as e:
            logging.error(f"Export error: {e}")
            messagebox.showerror("내보내기 실패", f"{e}")
    
    def filter_grid(self, event=None):
        logging.info("filter_grid 호출됨.")
        
        df_to_show = None 

        if self.current_step < 2:
            df_source = self.processor.aggregated_plan_df
            if df_source is not None and self.processor.date_cols:
                df_to_show = df_source[df_source[self.processor.date_cols].sum(axis=1) > 0].copy()
            else:
                df_to_show = df_source
        else:
            df_source = self.processor.simulated_plan_df
            if df_source is not None:
                shipment_cols = [col for col in df_source.columns if isinstance(col, str) and col.startswith('출고_')]
                if shipment_cols:
                    df_to_show = df_source[df_source[shipment_cols].sum(axis=1) > 0].copy()
                else:
                    df_to_show = df_source.copy()
            else:
                df_to_show = None

        if df_to_show is None:
            self.populate_master_grid(pd.DataFrame())
            return

        search_term = self.search_entry.get().lower()
        if search_term:
            df_to_show_reset = df_to_show.reset_index()
            df_to_show = df_to_show[df_to_show_reset['Model'].str.lower().str.contains(search_term).values]
            logging.info(f"검색어 '{search_term}' 필터링 후 크기: {df_to_show.shape}")
            
        self.populate_master_grid(df_to_show)

    def update_status_bar(self, message="준비 완료"):
        self.status_bar.configure(text=f"현재 파일: {self.current_file} | 상태: {message}")
        logging.info(f"상태 업데이트: {message}")

    def load_settings_to_gui(self):
        for key, entry_widget in self.settings_entries.items():
            entry_widget.delete(0, 'end')
            entry_widget.insert(0, str(self.config_manager.config.get(key, '')))
            
        for i, cb in self.day_checkboxes.items():
            if self.config_manager.config.get('DELIVERY_DAYS', {}).get(str(i), 'False') == 'True':
                cb.select()
            else:
                cb.deselect()
        logging.info("UI에 설정값 로드 완료.")

    def save_settings_and_recalculate(self):
        logging.info("설정 저장 및 재계산 시작.")
        new_config = self.config_manager.config.copy()
        try:
            for key, entry_widget in self.settings_entries.items():
                new_config[key] = int(entry_widget.get())
            
            new_delivery_days = {str(i): str(self.day_checkboxes[i].get()) for i in range(7)}
            new_config['DELIVERY_DAYS'] = new_delivery_days
            
            self.config_manager.save_config(new_config)
            self.processor.config = new_config
            
            if self.current_step >= 1 and self.processor.current_filepath:
                self.processor.process_plan_file()
            if self.current_step >= 2:
                self.recalculate_with_fixed_values()
                self.check_shipment_capacity()
                
            self.filter_grid()
            messagebox.showinfo("성공", "설정이 저장되었고 현재 단계까지 재계산되었습니다.")
            logging.info("설정 저장 및 재계산 완료.")
        except Exception as e:
            logging.error(f"Settings save and recalc error: {e}")
            messagebox.showerror("오류", f"설정 저장 및 재계산 실패: {e}")

    def open_daily_truck_dialog(self):
        logging.info("일자별 최대 차수 설정 다이얼로그 열기.")
        dialog = DailyTruckDialog(self, self.config_manager.config.get('DAILY_TRUCK_OVERRIDES', {}))
        self.wait_window(dialog)
        if dialog.result is not None:
            self.config_manager.config['DAILY_TRUCK_OVERRIDES'] = dialog.result
            self.save_settings_and_recalculate()
            logging.info("일자별 최대 차수 설정이 저장되었습니다.")

    def open_holiday_dialog(self):
        logging.info("휴무일/공휴일 설정 다이얼로그 열기.")
        current_holidays = [d for d in self.config_manager.config['NON_SHIPPING_DATES'] if isinstance(d, datetime.date)]
        dialog = HolidayDialog(self, current_holidays)
        self.wait_window(dialog)
        if dialog.result is not None:
            self.config_manager.config['NON_SHIPPING_DATES'] = dialog.result
            self.save_settings_and_recalculate()
            logging.info("휴무일/공휴일 설정이 저장되었습니다.")

    def open_safety_stock_dialog(self):
        logging.info("품목별 최소 재고 설정 다이얼로그 열기.")
        if self.processor.item_master_df is None:
            messagebox.showerror("오류", "품목 정보(Item.csv)가 로드되지 않았습니다.")
            return

        dialog = SafetyStockDialog(self, self.processor.item_master_df)
        self.wait_window(dialog)

        if dialog.result is not None:
            self.processor.item_master_df = dialog.result
            self.processor.save_item_master()
            messagebox.showinfo("저장 완료", "품목별 최소 재고 설정이 저장되었습니다. '설정 저장 및 재계산' 버튼을 눌러 계획에 반영하세요.")
            if self.current_step >=2:
                self.recalculate_with_fixed_values()


if __name__ == "__main__":
    try:
        app = ProductionPlannerApp(ConfigManager())
        app.mainloop()
    except Exception as e:
        logging.critical(f"Fatal error: {e}", exc_info=True)
        messagebox.showerror("치명적 오류", f"프로그램 실행에 실패했습니다.\n{e}")