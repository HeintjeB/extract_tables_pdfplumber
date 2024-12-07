import pdfplumber
from pdfplumber.table import Table, Row, Column
import pandas as pd
import traceback
import os 
file_path = os.path.dirname(__file__)
parent = os.path.dirname(file_path)
import re


class PurchaseOrderReader:
    def __init__(self, pdf):
        self.pdf = pdf
        self.pdf_shortcut = pdf.replace('.pdf', '')
        
    def pdfplumber_data_extractor(self):
        with pdfplumber.open(self.pdf) as self.doc:
            self.words_dict = {f'page {page}' : [self.doc.pages[page].extract_words()] for page in range(len(self.doc.pages))}
            self.pdfplumber_text = {f'page {page}' : [repr(self.doc.pages[page].extract_text())] for page in range(len(self.doc.pages))}
            self.pdf_extract_dict = {f'page {page}' : self.doc.pages[page] for page in range(len(self.doc.pages))}
        
    def draw_tables_for_determining(self):
        
        for page in self.pdf_extract_dict:
            with open(f'txt/coordinates_initial_{page}.txt', 'w') as file:
                file.write('')
            if not os.path.exists(f'txt/coordinates_{page}.txt'):
                with open(f'txt/coordinates_{page}.txt', 'w') as file:
                    file.write('')
            im_init = self.pdf_extract_dict[page].to_image()
            self.tables = self.pdf_extract_dict[page].find_tables()
            for idx, table in enumerate(self.tables):
                im_init.draw_rect(table.bbox, stroke='green')
                with open(f'txt/coordinates_initial_{page}.txt', 'a') as file:
                    tpl_str = str(idx)+":"+','.join(map(str, table.bbox,))
                    file.write(str(tpl_str) + '\n')
            im_init.save(f"png/determine_tables_for_{os.path.basename(self.pdf_shortcut)}_{page}.png")

    def creating_dataframe_based_on_coordinates(self):
        try:
            for pagenr, page in enumerate(self.pdf_extract_dict):
                coordinates = [tuple(map(float, line.split(':')[1].split(','))) for line in open(f"txt/coordinates_{page}.txt", "r")]
                tables_dict = {f'table {idx}' : Table(pagenr, [coordinate]) for idx, coordinate in enumerate(coordinates)}
                extract_dict = {table : repr(tables_dict[table].extract_2(page = self.doc.pages[tables_dict[table].page])[0][0]) for table in tables_dict}
                splitted_dict = {table: [re.split(r'(\s[0-9]{1,6})', i.replace("'", "")) for i in extract_dict[table].split('\\n')] for table in extract_dict}
                cleaned_dict = {table: [[item.strip() for item in lst if item.strip()] for lst in splitted_dict[table]] for table in splitted_dict}
                dataframe_dict = {table: pd.DataFrame(cleaned_dict[table][1:], columns=cleaned_dict[table][0]) for table in cleaned_dict}
                totaal_dataframe = pd.concat([dataframe_dict[table] for table in dataframe_dict], axis=1).T.drop_duplicates().T
                with pd.ExcelWriter(f'xlsx/totaal_dataframe_{page}.xlsx', engine='openpyxl') as writer:
                    totaal_dataframe.to_excel(writer, index=False, sheet_name='Overview')
                im = self.pdf_extract_dict[page].to_image()
                for table in tables_dict:
                    im.draw_rect(tables_dict[table].bbox, stroke='green')
                im.save(f"png/finetune_tables_for_{os.path.basename(self.pdf_shortcut)}_{page}.png")

        except:
            print(traceback.format_exc())

    def creating_dataframe_based_on_words(self):
        for page in self.words_dict:
            x_dict = {year : [word['x0'] for word in self.words_dict[page][0] if (re.match(fr'({year}\d{{2}})', str(word['text'])))] for year in ['2024','2025']}
            subcategory_list = [word['x0'] for word in self.words_dict[page][0] if 'Subcategory' in word['text']]
            for idx, table in enumerate(x_dict):
                x_dict[table].append(max([word['x1'] for word in self.words_dict[page][0] if re.match(fr'({table}\d{{2}})', str(word['text']))]))
                x_dict[table].append(subcategory_list[idx])
            y_list = list(set([word['top'] for word in self.words_dict[page][0] if re.match(r'(\D{1,10})', str(word['text']))]))
            y_list.append(max([word['bottom'] for word in self.words_dict[page][0] if re.match(r'(\D{1,10})', str(word['text']))]))
            dataframe_dict_2 = {f'table {idx}' : self.pdf_extract_dict[page].extract_table(table_settings={"vertical_strategy" : "explicit",
                                                    "horizontal_strategy" : "explicit",
                                    "explicit_vertical_lines": x_dict[table],
                                    "explicit_horizontal_lines" : y_list}) for idx, table in enumerate(x_dict)}
            totaal_dataframe_2 = pd.concat([pd.DataFrame(dataframe_dict_2[table]) for table in dataframe_dict_2], axis=1).T.drop_duplicates().T
            totaal_dataframe_2 = pd.DataFrame(totaal_dataframe_2.iloc[1:].values, columns=totaal_dataframe_2.iloc[0])
            with pd.ExcelWriter(f'xlsx/totaal_dataframe_2_{page}.xlsx', engine='openpyxl') as writer:
                totaal_dataframe_2.to_excel(writer, index=False, sheet_name='Overview')

    def creating_png_based_on_rows_columns(self):
        for page in self.pdf_extract_dict:
            im = self.pdf_extract_dict[page].to_image()
            path = f"png/columns_rows_for_{os.path.basename(self.pdf_shortcut)}.png"
            for table in self.tables:
                for i in table._get_rows_or_cols(kind=Row):
                    im.draw_rect(i.bbox, stroke='green')
                for i in table._get_rows_or_cols(kind=Column):
                    im.draw_rect(i.bbox, stroke='green')
            im.save(path)
        
if __name__ == '__main__':
    reader = PurchaseOrderReader(fr"{file_path}\pdf\transport_matrix.pdf")
    reader.pdfplumber_data_extractor()
    reader.draw_tables_for_determining()
    reader.creating_dataframe_based_on_coordinates()
    reader.creating_dataframe_based_on_words()
    reader.creating_png_based_on_rows_columns()