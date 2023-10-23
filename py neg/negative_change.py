import os
import pandas as pd
import numpy as np

#pd.io.formats.excel.ExcelFormatter.header_style = None


class Report:
    def __init__(self, now_file: str, prev_file: str, processes: list):

        now_name = self.get_name(now_file)
        prev_name = self.get_name(prev_file)

        print('Получение данных')
        now_df = pd.read_excel(now_file)
        prev_df = pd.read_excel(prev_file)

        out_file = f'auto анализ_негатив_{now_name}.xlsx'
        # path = list(os.path.split(now_file))[:-1]
        # path += [out_file]
        # out_path = os.path.join(*path)

        out_path = os.path.join("..", out_file)

        self.writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
        workbook = self.writer.book

        self.int_frmt = workbook.add_format({'align': 'right', 'num_format': '#'})
        self.perc_frmt = workbook.add_format({'align': 'right', 'num_format': '0%'})
        self.head_frmt = workbook.add_format({'align': 'center', 'bold': True})

        self.condition = {'type': '3_color_scale',
                          'min_type': 'min',
                          'min_color': '#5A8AC6',

                          'mid_type': 'percentile',
                          'mid_value': 50,
                          'mid_color': '#FCFCFF',

                          'max_type': 'max',
                          'max_color': '#F8696B'}

        for proc in processes:
            print(f'\tОбработка `{proc}`')

            if '+' in proc or '-' in proc:
                now = self.process_df(now_df, proc_name=proc, mode='now', sentiment=0)
                prev = self.process_df(prev_df, proc_name=proc, mode='prev', sentiment=0)

                proc_df = self.combine_df(now, prev)
                self.save(proc_df, proc_name=proc, now_name=now_name, prev_name=prev_name, sentiment=0)
            else:
                now_pos = self.process_df(now_df, proc_name=proc, mode='now', sentiment=1)
                prev_pos = self.process_df(prev_df, proc_name=proc, mode='prev', sentiment=1)
                now_neg = self.process_df(now_df, proc_name=proc, mode='now', sentiment=-1)
                prev_neg = self.process_df(prev_df, proc_name=proc, mode='prev', sentiment=-1)

                proc_df_pos = self.combine_df(now_pos, prev_pos)
                proc_df_neg = self.combine_df(now_neg, prev_neg)
                self.save(proc_df_pos, proc_name=proc, now_name=now_name, prev_name=prev_name, sentiment=1)
                self.save(proc_df_neg, proc_name=proc, now_name=now_name, prev_name=prev_name, sentiment=-1)

        print('сохранение')
        self.writer.close()
        #self.writer.save()
        print(f'сохранено в {out_path}')

    @staticmethod
    def get_name(filepath: str) -> str:
        filename = filepath.split('/')[-1]
        name = filename.split('_')[0]
        return name

    @staticmethod
    def process_df(df: pd.DataFrame, proc_name: str, mode: str, sentiment) -> pd.Series:
        out = df[df['Process'] == proc_name]['Name'].value_counts()
        out.index.name = 'Причина'
        out.name = mode

        def drop_indexes(out, sentiment):
            out_copy = out.copy()
            for i in range(len(out_copy)):
                sen = df[df['Name'] == out.index[i]]['Sentiment'].values
                if sen[0] == -1 * sentiment:
                    out_copy = out_copy.drop(out.index[i])
                if sentiment == 1 and sen[0] == 0:
                    out_copy = out_copy.drop(out.index[i])
            return out_copy

        if sentiment == 1:
            out_pos = drop_indexes(out, sentiment)
            return out_pos.copy()
        elif sentiment == -1:
            out_neg = drop_indexes(out, sentiment)
            return out_neg.copy()
        else:
            return out.copy()


    @staticmethod
    def combine_df(s1: pd.Series, s2: pd.Series) -> pd.DataFrame:
        df = pd.concat([s1, s2], axis=1)
        df.fillna(0, inplace=True)
        for col in ['now', 'prev']:
            df[col] = df[col].astype(int)
        df['прирост, абс.'] = df['now'] - df['prev']
        df['прирост, %'] = df['прирост, абс.'] / df['prev']
        df['прирост, %'] = df['прирост, %'].replace(np.inf, np.nan)
        return df

    def save(self, df: pd.DataFrame, proc_name: str, now_name: str, prev_name: str, sentiment) -> None:

        if sentiment == 1:
            proc_name = proc_name + '+'
        elif sentiment == -1:
            proc_name = proc_name + '-'
        else:
            pass

        if '/' in proc_name:
            proc_name = proc_name.replace('/', ' или ')

        df.reset_index(inplace=True)
        df.rename(columns={'now': now_name, 'prev': prev_name}, inplace=True)
        df.to_excel(self.writer, index=False, sheet_name=proc_name)

        worksheet = self.writer.sheets[proc_name]
        worksheet.set_column('A:A', width=60)
        worksheet.set_column('B:C', width=15, cell_format=self.int_frmt)
        worksheet.set_column('D:D', width=15)
        worksheet.set_column('E:E', width=15, cell_format=self.perc_frmt)
        worksheet.set_row(0, cell_format=self.head_frmt)

        end = len(df) + 1
        worksheet.conditional_format(f"D1:D{end}", self.condition)
        worksheet.conditional_format(f"E1:E{end}", self.condition)

# if __name__ == '__main__':
#     Report(now_file= './30.05-05.06_остальное.xlsx',
#            prev_file= './23-29.05_остальное.xlsx',
#            processes= ['Питание', 'Борт',
#                        'Регулярность', 'Регистрация',
#                        'Цены', 'Напитки',
#                        'БП -'])