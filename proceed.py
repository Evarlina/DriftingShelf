import json
from re import findall
from string import ascii_lowercase, ascii_uppercase

import xlwings as xw


class SheetOperation:

    def __init__(self):
        """ Create class variables.
        """
        # Open the working sheets.
        print('打开表格...', end='  ')
        self.app = xw.App(visible=True, add_book=False)
        # app.display_alerts = False
        # app.screen_updating = False
        self.wb = self.app.books.open('./漂流书架.xlsx')

        # Open bookinfo sheet.
        self.bookinfo_sheet = self.wb.sheets[0]
        self.bookinfo_sheet.range('A2').expand().api.Delete()

        # Open proof sheet.
        self.proof_sheet = self.wb.sheets[1]

        # Dictionary where category info is stored.
        self.category_dic = {}

    def load_json(self):
        """ Load json file.
        """
        print('加载json...', end='  ')
        with open('./category.json', 'r', encoding='utf-8') as fr:
            self.category_dic = json.load(fr)

    def alert_on_quit(self):
        """ Remind user before quit too early.
        """
        # HCI part.
        print('-' * 40)
        will_quit = input('提前退出会导致当前行信息丢失，确定退出？(y/n)\n>> ').strip()

        # If the input is neither 'y' nor 'n', re-input.
        while will_quit not in ['y', 'n']:
            will_quit = input('输入错误，请重新选择！\n>> ').strip()

        # If about to quit, save the workbook, and exit with code 1.
        if will_quit == 'y':
            self.close()
            exit(1)

    def handle_title(self):
        """ Get title info.
        """
        # HCI part.
        print('-' * 40)
        title = input('请输入书籍名（不带书名号）：\n>> ').strip()
        while title == '':
            title = input('>> ').strip()

        # If user wants to quit, give an alert.
        if title == 'quit':
            self.alert_on_quit()

        # Process input.
        title = '《' + title + '》'
        return title

    def handle_author(self):
        """ Get author info.
        """
        # HCI part.
        print('-' * 40)
        author_list = input('请输入作者名，多个作者以空格分隔：\n>> ').strip()
        while author_list == '':
            author_list = input('>> ').strip()

        # If user wants to quit, give an alert.
        if author_list == 'quit':
            self.alert_on_quit()

        # Process input.
        author_list = author_list.split()
        author = '、'.join(author_list)
        return author

    def handle_identifier(self):
        """ Get category and identifier info.
        """
        # HCI part.

        # Print lvl1 categories currently available.
        print('-' * 40)
        print('请选择一级类别：')
        for item in list(self.category_dic.keys()):
            print(item, end='\t')

        # Get chief category from user.
        chief_cat = input('\n\n>> ').strip()
        while chief_cat == '':
            chief_cat = input('>> ').strip()
        if chief_cat == 'quit':
            self.alert_on_quit()
        while chief_cat not in [str(i) for i in range(1, 4)]:
            chief_cat = input('输入错误，请重新选择！\n>> ').strip()

        # Translate it into uppercase letter, and add to identifier.
        chief_cat = int(chief_cat)
        upper_alphabet = list(ascii_uppercase)[:3]
        identifier = upper_alphabet[chief_cat - 1]

        # HCI part.

        # Print lvl2 categories currently available.
        print('-' * 40)
        print('请选择二级类别：')
        detail_list = self.category_dic[list(enumerate(self.category_dic))[
            chief_cat - 1][1]]
        for item in detail_list:
            print(item)

        # Get second category from user.
        second_cat = input('\n>> ').strip()
        while second_cat == '':
            second_cat = input('>> ').strip()
        if second_cat == 'quit':
            self.alert_on_quit()
        while second_cat not in [str(i) for i in range(1, len(detail_list) + 1)]:
            second_cat = input('输入错误，请重新选择！\n>> ').strip()

        # Translate it into lowercase letter, and add to identifier.
        second_cat = int(second_cat)
        lower_alphabet = list(ascii_lowercase)[:16]
        identifier += lower_alphabet[second_cat - 1]

        # Beautify the category Chinese name.
        cat_fin = findall(r'\d+::(.*)', detail_list[second_cat - 1])[0]

        # HCI part.
        print('-' * 40)
        serial_no = input(f'该书的编号前缀为{identifier}，请输入它在该类中的数字编号：\n>> ').strip()
        while serial_no == '':
            serial_no = input('>> ').strip()
        if serial_no == 'quit':
            self.alert_on_quit()
        identifier += '-' + serial_no.zfill(2)

        return cat_fin, identifier

    def handle_donor(self):
        """ Get donor info.
        """
        # HCI part.
        print('-' * 40)
        donor = input('请输入捐赠者姓名：\n>> ').strip()
        while donor == '':
            donor = input('>> ').strip()

        # If user wants to quit, give an alert.
        if donor == 'quit':
            self.alert_on_quit()
        return donor

    def handle_comment(self):
        """ Get comment info.
        """
        # HCI part.
        print('-' * 40)
        comment = input('请输入捐赠留言：\n>> ').strip()
        while comment == '':
            comment = input('>> ').strip()

        # If user wants to quit, give an alert.
        if comment == 'quit':
            self.alert_on_quit()
        return comment

    def start_loop(self):
        """ Main loop where data is get and stored.
        """
        # State certain variables.
        bookinfo_cursor = 2
        proof_cursor = self.proof_sheet.range(
            'A1').expand('down').last_cell.row + 1
        donor = ''

        # Start loop.
        print('开始读入信息。')
        quit = False
        while not quit:

            # Get book data.
            title = self.handle_title()
            author = self.handle_author()
            category, bookid = self.handle_identifier()
            pre_donor = donor
            donor = self.handle_donor()
            comment = self.handle_comment()

            # Add to bookinfo sheet.
            self.bookinfo_sheet.range(f'A{bookinfo_cursor}').value = [
                title, author, category, bookid, donor, comment]

            # Beautify the sheet.
            self.bookinfo_sheet.autofit()

            # Add to proof sheet.
            proofid = str(proof_cursor - 1).zfill(3)

            # If donor name repeats.
            if donor == pre_donor:
                print('-' * 40)
                will_cover = input(
                    f'捐赠证明提示：检测到前一本书的捐赠者也是{donor}，是否合并？(y/n)\n>> ').strip()
                while will_cover not in ['y', 'n']:
                    will_cover = input(f'输入错误，请重新选择！(y/n)\n>> ').strip()

                # If yes, move up cursor and merge.
                if will_cover == 'y':
                    proof_cursor -= 1
                    self.proof_sheet.range(
                        f'B{proof_cursor}').value += f'\n{bookid}'

                # If no, persistence continues.
                else:
                    self.proof_sheet.range(f'A{proof_cursor}').value = [
                        proofid, bookid]

            # If not repeated.
            else:
                self.proof_sheet.range(f'A{proof_cursor}').value = [
                    proofid, bookid]

            # Beautify the sheet.
            self.proof_sheet.autofit()

            # Determine whether to continue.
            print('-' * 40)
            choice = input('信息添加完毕！您现在可以输入"quit"以退出，或其他字符以继续。\n>> ').strip()
            if choice == 'quit':
                quit = True
            else:
                bookinfo_cursor += 1
                proof_cursor += 1

    def close(self):
        """ To save and quit program.
        """
        self.bookinfo_sheet.autofit()
        self.proof_sheet.autofit()
        self.wb.save()
        self.wb.close()
        self.app.quit()


if __name__ == '__main__':
    opr = SheetOperation()
    opr.load_json()
    opr.start_loop()
    opr.close()
