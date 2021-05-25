from json import load
from re import findall
from string import ascii_lowercase, ascii_uppercase
from sys import exit

import xlwings as xw

from query_data import get_data


class SheetOperation:

    def __init__(self):
        """ Create class variables.
        """
        # Create app and configure.
        print('打开表格...', end='  ')
        self.app = xw.App(visible=True, add_book=False)
        # self.app.display_alerts = False
        # self.app.screen_updating = False

        # Open "书籍信息.xlsx".
        try:
            self.bookinfo_wb = self.app.books.open('./rsc/书籍信息.xlsx')
        except FileNotFoundError:
            print('\n找不到文件"书籍信息.xlsx"，请确认它是否位于/rsc目录下。')

        # Open "捐赠证明.xlsx".
        try:
            self.proof_wb = self.app.books.open('./rsc/捐赠证明.xlsx')
        except FileNotFoundError:
            print('\n找不到文件"捐赠证明.xlsx"，请确认它是否位于/rsc目录下。')

        # Open bookinfo sheet and clear.
        self.bookinfo_sheet = self.bookinfo_wb.sheets[0]
        self.bookinfo_sheet.range('A2').expand().api.Delete()

        # Open proof sheet and clear.
        self.proof_sheet = self.proof_wb.sheets[0]
        self.proof_sheet.range('A2').expand().api.Delete()
        self.proof_sheet.range('A:A').api.NumberFormat = '@'

        # Dictionary where category info is stored.
        self.category_dic = {}

    def load_json(self):
        """ Load "category.json".
        """
        print('加载json...', end='  ')
        try:
            with open('./rsc/category.json', 'r', encoding='utf-8') as fr:
                self.category_dic = load(fr)
        except FileNotFoundError:
            print('\n找不到文件"category.json"，请确认它是否位于/rsc目录下。')

    def alert_on_quit(self, question):
        """ Remind user before quit too early.
        """
        # HCI part.
        print('-' * 60)
        will_quit = input('提前退出会导致当前行的信息丢失，确定退出？(y/n)\n>> ').strip()

        # If no response, keep asking.
        while will_quit == '':
            will_quit = input('>> ').strip()

        # If the input is neither 'y' nor 'n', re-input.
        while will_quit not in ['y', 'n']:
            will_quit = input('输入错误，请重新选择！\n>> ').strip()

        # If yes, save the workbook and exit with code 1.
        if will_quit == 'y':
            self.close()
            exit(1)

        # If no, continue to input the answer to the former question.
        else:
            print('-' * 60)
            print(f'请继续输入刚才的{question}：', end='')

            # If the question were about donor and comment, hint user that he could use 'y' to auto-fill.
            if question == '捐赠者姓名' and self.pre_donor != '':
                print(f'键入"y"可填入上一位捐赠者的姓名（{self.pre_donor}）。', end='')
            elif question == '捐赠留言' and self.pre_comment != '':
                print(f'键入"y"可填入上一条留言（{self.pre_comment}）。', end='')

            regret = input('\n>> ').strip()
            return regret

    def handle_title(self):
        """ Get title info.
        """
        # HCI part.
        print('-' * 60)
        title = input('请输入书籍名：\n>> ').strip()

        # If no response, keep asking.
        while title == '':
            title = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if title == 'quit':
            title = self.alert_on_quit('书籍名')

        print('-' * 60)
        print('正在豆瓣读书网查询书籍信息......', end='')
        self.recommended_authors, self.recommended_tags = get_data(title)
        print('成功！')

        # Process input.
        if title.find('《') == -1:
            title = '《' + title
        if title.find('》') == -1:
            title = title + '》'

        return title

    def handle_author(self):
        """ Get author info.
        """
        # HCI part.
        print('-' * 60)
        print('找到了下列可能的作者：')
        for item in self.recommended_authors.items():
            print(item[0] + '-' * 10 + item[1])
        author_list = input('\n请输入作者名，多个作者以空格分隔：\n>> ').strip()

        # If no response, keep asking.
        while author_list == '':
            author_list = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if author_list == 'quit':
            author_list = self.alert_on_quit('作者名')

        # Process input.
        author_list = author_list.split()
        author = '、'.join(author_list)

        return author

    def handle_identifier(self):
        """ Get category and identifier info.
        """
        # HCI part.
        print('-' * 60)
        print('本书最热门的五个标签：')
        print('、'.join(self.recommended_tags))
        print('\n请选择一级类别：')
        for item in list(self.category_dic.keys()):
            print(item, end='\t')
        chief_cat = input('\n\n>> ').strip()

        # If no response, keep asking.
        while chief_cat == '':
            chief_cat = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if chief_cat == 'quit':
            chief_cat = self.alert_on_quit('一级类别')

        # If the input is not in ['1', '2', '3'], re-input.
        while chief_cat not in [str(i) for i in range(1, 4)]:
            chief_cat = input('输入错误，请重新选择！\n>> ').strip()

        # Translate it into uppercase letter, and add to identifier.
        chief_cat = int(chief_cat)
        upper_alphabet = list(ascii_uppercase)[:3]
        identifier = upper_alphabet[chief_cat - 1]

        # HCI part.
        print('-' * 60)
        print('请选择二级类别：')
        detail_list = self.category_dic[list(enumerate(self.category_dic))[
            chief_cat - 1][1]]  # Locate the detail list of lvl1 category.
        for item in detail_list:
            print(item)
        second_cat = input('\n>> ').strip()

        # If no response, keep asking.
        while second_cat == '':
            second_cat = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if second_cat == 'quit':
            second_cat = self.alert_on_quit('二级标题')

        # If the input is not in the range of the detail list, re-input.
        while second_cat not in [str(i) for i in range(1, len(detail_list) + 1)]:
            second_cat = input('输入错误，请重新选择！\n>> ').strip()

        # Translate it into lowercase letter, and add to identifier.
        second_cat = int(second_cat)
        lower_alphabet = list(ascii_lowercase)[:16]
        identifier += lower_alphabet[second_cat - 1]

        # Beautify the category's Chinese name.
        cat_cn = findall(r'\d+::(.*)', detail_list[second_cat - 1])[0]

        # HCI part.
        print('-' * 60)
        serial_no = input(
            f'该书类别为"{cat_cn}"({identifier})，请输入它在该类中的数字编号：\n>> ').strip()

        # If no response, keep asking.
        while serial_no == '':
            serial_no = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if serial_no == 'quit':
            serial_no = self.alert_on_quit('数字编号')

        # Process input and add to identifier.
        identifier += '-' + serial_no.zfill(2)

        return cat_cn, identifier

    def handle_donor(self):
        """ Get donor info.
        """
        # HCI part.
        print('-' * 60)
        print('请输入捐赠者姓名：', end='')
        if self.pre_donor != '':
            print(f'键入"y"可填入上一位捐赠者的姓名（{self.pre_donor}）。', end='')
        donor = input('\n>> ').strip()

        # If no response, keep asking.
        while donor == '':
            donor = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if donor == 'quit':
            donor = self.alert_on_quit('捐赠者姓名')

        # If the donor is the same, assign that.
        if donor == 'y':
            donor = self.pre_donor

        return donor

    def handle_comment(self):
        """ Get comment info.
        """
        # HCI part.
        print('-' * 60)
        print('请输入捐赠留言：', end='')
        if self.pre_comment != '':
            print(f'键入"y"可填入上一条留言（{self.pre_comment}）。', end='')
        comment = input('\n>> ').strip()

        # If no response, keep asking.
        while comment == '':
            comment = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if comment == 'quit':
            comment = self.alert_on_quit('捐赠留言')

        # If comment is the same, assign that.
        if comment == 'y':
            comment = self.pre_comment

        return comment

    def start_loop(self):
        """ Main loop where data is get and stored.
        """
        # State certain variables.
        bookinfo_cursor = 2
        proof_cursor = 2
        donor = ''
        comment = ''

        # Start to read in data.
        print('开始读入信息。')

        # Determine starting index of proof.
        print('-' * 60)
        proof_delta = input('请输入本次的起始捐赠证明编号：\n>> ').strip()

        # If no response, keep asking.
        while proof_delta == '':
            proof_delta = input('>> ').strip()

        # If user wants to quit, raise an alert.
        if proof_delta == 'quit':
            self.alert_on_quit()

        # Process input.
        proof_delta = int(findall('0*(\d+)', proof_delta)[0]) - proof_cursor

        # Start loop.
        quit = False
        while not quit:

            # Get book data.
            title = self.handle_title()
            author = self.handle_author()
            category, bookid = self.handle_identifier()
            self.pre_donor = donor
            donor = self.handle_donor()
            self.pre_comment = comment
            comment = self.handle_comment()

            # Add to bookinfo sheet.
            self.bookinfo_sheet.range(f'A{bookinfo_cursor}').value = [
                title, author, category, bookid, donor, comment]

            # Beautify bookinfo sheet.
            self.bookinfo_sheet.autofit()

            # Calculate current ID of proof.
            proofid = str(proof_cursor + proof_delta).zfill(3)

            # If donor's name repeats.
            if donor == self.pre_donor:

                # HCI part.
                print('-' * 60)
                print(f'检测到前一本书的捐赠者也是{donor}，是否合并捐赠信息？(y/n)')
                print('提示：如果刚刚撤销过该捐赠者的第一条捐赠，也请选择"y"。')
                will_merge = input('>> ').strip()

                # If no response, keep asking.
                while will_merge == '':
                    will_merge = input('>> ').strip()

                # If the input is neither 'y' nor 'n', re-input.
                while will_merge not in ['y', 'n']:
                    will_merge = input(f'输入错误，请重新选择！(y/n)\n>> ').strip()

                # If yes, move up cursor and merge.
                if will_merge == 'y':
                    proof_cursor -= 1

                    # Calculate current ID of proof.
                    proofid = str(proof_cursor + proof_delta).zfill(3)
                    self.proof_sheet.range('A:A').api.NumberFormat = '@'

                    # Check whether this row has identifier.
                    if self.proof_sheet.range(f'B{proof_cursor}').value is None:
                        self.proof_sheet.range(f'A{proof_cursor}').value = [
                            proofid, bookid]
                    else:
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

            # Beautify proof sheet.
            self.proof_sheet.autofit()

            # Determine whether to continue.
            print('-' * 60)
            choice = input(
                '信息添加完毕！您现在可以：\n- 输入"quit"以退出；\n- 输入"undo"以撤销本行操作；\n- 输入其他任意字符以继续添加信息。\n>> ').strip()

            # Quit loop if agreed.
            if choice == 'quit':
                quit = True

            # Undo this line if agreed.
            elif choice == 'undo':

                # Delete this line in bookinfo sheet.
                self.bookinfo_sheet.range(
                    f'A{bookinfo_cursor}').expand('right').api.Delete()

                # If the current donor has multiple identifier, delete the last one only.
                identifier_list = self.proof_sheet.range(
                    f'B{proof_cursor}').value

                if '\n' in identifier_list:
                    identifier_list = '\n'.join(
                        identifier_list.split('\n')[:-1])
                    self.proof_sheet.range(
                        f'B{proof_cursor}').value = identifier_list

                # If the donor is a new one, delete the whole row.
                else:
                    self.proof_sheet.range(f'A{proof_cursor}').expand(
                        'right').api.Delete()

                proof_cursor += 1

            # If about to continue, move down the cursors and restart the loop.
            else:
                bookinfo_cursor += 1
                proof_cursor += 1

    def close(self):
        """ To save and quit program.
        """
        # Close bookinfo sheet.
        self.bookinfo_sheet.autofit()
        self.bookinfo_wb.save()
        self.bookinfo_wb.close()

        # Close proof sheet.
        self.proof_sheet.autofit()
        self.proof_wb.save()
        self.proof_wb.close()

        # Quit app.
        self.app.quit()


if __name__ == '__main__':
    opr = SheetOperation()
    opr.load_json()
    opr.start_loop()
    opr.close()
