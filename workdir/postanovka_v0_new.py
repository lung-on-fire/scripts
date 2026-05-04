import sys, glob, os
import pandas as pd
import time
import warnings
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Font, Alignment

def main():
    try:
        files = glob.glob('*.xlsx')
        if len(files) != 1:
            raise ValueError(f"Найдено {len(files)} XLSX файлов. Требуется ровно один файл.")
        filename = files[0]

        # Называет выходной файл в зависимости от времени системы
        if (int(time.strftime("%H")) > 12) or (int(time.strftime("%H")) < 6):
            output_file = f"../postanovka_{time.strftime('%d%m%y')}_ночь_v0.xlsx"
        else:
            output_file = f"../postanovka_{time.strftime('%d%m%y')}_день_v0.xlsx"

        # Здесь можно указать номер листа (по умолчанию первый = 0)
        data = pd.read_excel(filename, sheet_name=0)
        data.iloc[:, 4] = data.iloc[:, 4].astype(str).str.strip()

        # Заполнение первого столбика сплошь номерами
        data.iloc[:, 0] = data.iloc[:, 0].ffill()

        if os.path.exists('otchet.txt'):
            os.remove('otchet.txt')

        # Получаем список инфекций
        # CAV = CAVI + CAVII
        # HV = CHV + FHV
        # PV = CPV + FPV
        infections = ['FCV', 'CPIV', 'FCoV', 'CCoV', 'СCoV', 'CDV', 'Нью', 'Giard', 'Salm', 'Tritri', 'Bartonella', 'Токсопл', 'gibsoni',
                      'Асперг', 'Орнитоз', 'Полио', 'Цирко','Bruc', 'haemofelis', 'haеmocanis', 'perfringens', 'galiseptica', 'Паст', 'B.canis', 'Babesia canis',
                      'B.spp', 'Babesia spp', 'РНК', 'HV', 'Мycoplasma spp', 'Mycoplasma spp', 'M.felis','Mycoplasma felis', 'FIV', 'FeLV', 'PV', 'Campylobacter spp', 'Campilobacter jejuni', 'Клостр', 'Борд', 'Chlamyd', 
                      'Crypto', 'CAV', 'M.canis', 'Anaplasma', 'Борр', 'Ehr', 'Urea', 'Диро', 'Lepto', 'Microsp', 'Trich', 'Bruc',
                      'Haemobartonella felis', 'Haemobartonella canis', 'Бабезиоз собак и кошек']
        
        contract_infections = ['Haemobartonella felis', 'Haemobartonella canis', 'Бабезиоз собак и кошек']
        

        results = {}
        bcanis_nums1 = []
        bcanis_nums2 = []
        gibsoni_nums  = []
        bspp_numbers1  = []
        bspp_numbers2  = []
        bspp_imposters = []
        felis_nums = []
        felis_new = []


        def normalize_keys(key):
            key = key.replace('М', 'M') 
            key = key.replace('С', 'C')
            return key

        for infection in infections:
            result_key = normalize_keys(infection.strip())

            if (result_key == "Mycoplasma spp"):
                pattern = r'[MМ]ycoplasma\s+spp'
                infection_data = data[data.iloc[:, 4].str.contains(pattern, case=True, na=False, regex=True)]

            elif (result_key == "CCoV"):
                pattern = r'[CС]oV'
                mask_c = data.iloc[:, 4].str.contains(pattern, case=True, na=False, regex=True)
                mask_f = data.iloc[:, 4].str.contains('FCoV', case=True, na=False, regex=False)
                infection_data = data[mask_c & ~mask_f]
            
            elif (result_key == "FCoV"):
                mask_f = data.iloc[:, 4].str.contains('FCoV', case=True, na=False, regex=False)
                infection_data = data[mask_f]

            else:
                infection_data = data[data.iloc[:, 4].str.contains(result_key, case=True, na=False, regex=False)]
            
            infection_numbers = infection_data.iloc[:, 0]
            results[result_key] = [int(p) for p in infection_numbers.tolist()]



        #  FIV + FeLV
        FIV_FeLV_numbers = set()
        FIV_numbers = set()
        FeLV_numbers = set()
        FeLV_RNA = set()


        for infection, numbers in results.items():
            if 'FIV' in results:
                FIV_only = [inf for inf in numbers if inf in results['FIV']]
                for number in FIV_only:
                    FIV_numbers.add(number)

            if 'FeLV' in results:
                FeLV_only = [inf for inf in numbers if inf in results['FeLV']]
                for number in FeLV_only:
                    FeLV_numbers.add(number)

            if ('FIV' and 'FeLV') in results:
                common_numbers = [inf for inf in numbers if inf in results['FIV'] and inf in results['FeLV']]
                for number in common_numbers:
                    FIV_FeLV_numbers.add(number)

            if 'РНК' in results:
                FeLV_RNA_num = [inf for inf in numbers if inf in results['РНК']]
                for number in FeLV_RNA_num:
                    FeLV_RNA.add(number)

        if 'Бабезиоз собак и кошек' in results:
            bspp_imposters = results.get('Бабезиоз собак и кошек', [])
        if 'B.canis' in results:
            bcanis_nums1 = results.get('B.canis', [])
        if 'Babesia canis' in results:
            bcanis_nums2 = results.get('Babesia canis', [])
        if 'Babesia spp' in results:
            bspp_numbers1 = results.get('Babesia spp', [])
        if 'B.spp' in results:
            bspp_numbers2 = results.get('B.spp', [])
        if 'gibsoni' in results:
            gibsoni_nums = results.get('gibsoni', [])


        if 'M.felis' in results:
            felis_nums = results.get('M.felis', [])
        if 'Mycoplasma felis' in results:
            felis_new = results.get('Mycoplasma felis', [])

        if felis_new:
            felis_nums = set(felis_nums + felis_new) 
        
        results['M.felis'] = [num for num in felis_nums]
        results['M.felis'] = sorted(results['M.felis'] )

        
        bcanis_nums = set(bcanis_nums1  + bcanis_nums2)
        bspp_numbers = set(bspp_numbers1 + bspp_numbers2 + bspp_imposters) 

        results['Babesia canis'] = [num for num in bcanis_nums]
        results['gibsoni'] = [num for num in gibsoni_nums]
        results['Babesia spp'] = [num for num in bspp_numbers]
        
        if bspp_imposters:
            #results.setdefault('Babesia spp', []).extend(bspp_numbers)
            results['Babesia spp'] = [num for num in bspp_numbers]
            results['Babesia spp'] = sorted(results['Babesia spp'])

            results['Babesia canis'] = [num for num in bcanis_nums if num not in bspp_imposters]
            results['Babesia canis'] = sorted(results['Babesia canis'])

            results['gibsoni'] = [num for num in gibsoni_nums if num not in bspp_imposters]
            results['gibsoni'] = sorted(results['gibsoni'])


        # Преобразуем множества в списки, удаляем старые FIV, FeLV и перезаписываем
        if (FIV_numbers | FeLV_numbers | FIV_FeLV_numbers):
            FIV_diff = FIV_numbers.difference(FIV_FeLV_numbers)
            FeLV_diff = FeLV_numbers.difference(FIV_FeLV_numbers)
            FIV_FeLV_numbers = sorted(list(FIV_FeLV_numbers))
            FIV_numbers = sorted(list(FIV_diff))
            FeLV_numbers = sorted(list(FeLV_diff))
            FeLV_RNA = sorted(list(FeLV_RNA))

            results['FIV+FeLV'] = FIV_FeLV_numbers
            results['FIV'] = FIV_numbers
            results['FeLV'] = FeLV_numbers
            if 'РНК' in results:
                results.pop('РНК')
                results['РНК FeLV'] = FeLV_RNA

        #with open('otchet.txt', 'a', encoding='utf-8') as file:
        #    if infection not in infections:
        #        file.write(f"Инфекция не поймана скриптом! Название и номера: {infection, numbers}")
        #    else:
        #        file.write("Новых инфекций-сюрпризов от клиник не обнаружено")

        # Группировка инфекций по категориям
        category_mapping = {
            'Заяц РНК': ['FCV', 'CPIV', 'FCoV', 'CCoV', 'CDV', 'Нью'],
            'Заяц ДНК': ['Giard', 'Salm', 'Tritri', 'Bartonella', 'Токсопл', 'Microsp','Trich', 'gibsoni', 'Орнитоз', 'Полио', 'Цирко'],
            'Genlab': ['HV', 'Mycoplasma spp', 'M.felis', 'Chlamyd', 'M.canis', 'FIV', 'FeLV', 'Campilobacter jejuni', 'Клостр', 'FIV+FeLV'],
            'Fractal': ['Борд', 'PV', 'Crypto', 'CAV', 'Anaplasma', 'Борр', 'Ehr', 'Urea', 'Диро', 'Lepto'],
            'VectBest': ['Bruc','haemofelis', 'haеmocanis', 'perfringens', 'galiseptica', 'Паст','Асперг','Babesia canis', 'Babesia spp', 'РНК FeLV', 'Haemobartonella felis', 'Haemobartonella canis', 'Campylobacter spp']
        }


        # Создаем новый DataFrame для вывода
        output_data = []

        for category, infections_in_category in category_mapping.items():
            output_data.append([category] + [''] * 10) 
            for infection in infections_in_category:
                infection = normalize_keys(infection)
                if infection in results:
                    numbers = results[infection]
                    output_data.append(['', infection] + numbers + [''] * (10 - len(numbers)))


        # Создаем Excel файл
        wb = Workbook()
        ws = wb.active

        thin_border = Border(
            left=Side(border_style="thin", color="000000"),
            right=Side(border_style="thin", color="000000"),
            top=Side(border_style="thin", color="000000"),
            bottom=Side(border_style="thin", color="000000")
        )

        thick_border = Border(
            left=Side(border_style="thick"),
            right=Side(border_style="thick"),
            top=Side(border_style="thick"),
            bottom=Side(border_style="thick")
        )

        # Записываем данные в Excel
        row_index = 1
        max_columns = max(len(row) for row in output_data)

        # Добавляем нумерацию и делаем нумерацию жирной
        column_numbers = ['', ''] + [str(i) for i in range(1, max_columns - 1)]
        ws.append(column_numbers)
        for col in range(1, max_columns + 1):
            cell = ws.cell(row=row_index, column=col)
            cell.border = thick_border
        row_index += 1

        # Записываем остальные данные
        for row in output_data:
            ws.append(row)
            if row[0]:  #Если это строка с категорией
                # Определяем количество строк для текущей категории
                category = row[0]
                infections_in_category = category_mapping.get(category, [])
                num_rows = len(infections_in_category)
                
                # Объединяем ячейки для заголовка категории вертикально
                ws.merge_cells(start_row=row_index, start_column=1, end_row=row_index + num_rows, end_column=1)
                
                # Центрируем текст в объединенной ячейке
                cell = ws.cell(row=row_index, column=1)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True)  # Жирный шрифт для заголовка категории
                
                # Применяем толстые границы для первого столбца (после объединения)
                for r in range(row_index, row_index + num_rows + 1):
                    cell = ws.cell(row=r, column=1)
                    cell.border = thick_border
            
            # Применяем жирный шрифт для подкатегорий (второй столбец)
            if row[1]:  #Если второй столбец не пустой (подкатегория)
                cell = ws.cell(row=row_index, column=2)
                cell.font = Font(bold=True)
            
            row_index += 1

        # Применяем стили границ
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border

        # Сохраняем файл
        wb.save(output_file)
        print(f"Результаты записаны в файл: {output_file}")

    except Exception as error:
        print(f"Error!{error}")
        sys.exit(1)

if __name__ == '__main__':
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")
    main()