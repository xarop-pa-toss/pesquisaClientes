# Programa de Pesquisa de Clientes. Ver 
# Utiliza PyQt5 para construção do UI e gestão segura das queries SQL. Liga-se a um ficheiro Access na mesma pasta que a app.
# ---TODO---
# -- Input do utilizador para escolher localização do ficheiro Access
# -- Mais segurança. Ecrã de login com password encriptada.
# -- Meter os Widgets num grupo para poderem ser movimentados/escalados em conjunto independentemente da resolução


from PyQt5.QtWidgets import QApplication, QPushButton, QComboBox, QLabel, QWidget, QTableWidget, QTableWidgetItem, QLineEdit, QMainWindow, QAction
from PyQt5.QtCore import QSize
from PyQt5.QtSql import QSql, QSqlDatabase, QSqlQuery
import sys

# Ligação pyodbc ao ficheiro Access. Variável 'cursor' usada para queries.
db = QSqlDatabase.addDatabase("QODBC")
db.setDatabaseName(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\STUFF\Python Projects\PesquisaClientes\DB_PesquisaClientes.accdb")
if db.open(): print("**** DB Connection established **** :)")


class MainWindow(QMainWindow):

    # 1 - Criação da janela principal que serve de base
    def __init__(self):
        super(MainWindow, self).__init__()
        
        self.setWindowTitle('Pesquisa de Clientes') # nome da janela principal
        self.setFixedSize(QSize(1125, 600)) # primeiros dois parâmetros são posição no ecrã. Terceiro e quarto parâmetros são largura x altura.

        self.add_menu_bar()
        self.add_search_items()
        self.add_table()

        label = QLabel('<h3>Procurar por:</h3>', parent = self)
        label.move(20, 30)
        
        self.queryButton.clicked.connect(self.execute_query)


    def execute_query(self):
        query = QSqlQuery()
        
        # Se nenhum texto for inserido na textbox, popula tabela com todos os resultados na tabela.
        # Se houver texto, procura pelo texto na coluna correspondente à combobox + textbox
        print(self.queryText.text())
        if self.queryText.text() == '*':
            query.prepare('SELECT * FROM Clientes')
        else:
            col = self.queryBox.currentText()
            val = self.queryText.text()
            print("col é: "+col)
            print("val é: "+val)
            if col in ('NIF', 'Telef','Telem'):
                query.prepare("SELECT * FROM Clientes WHERE "+col+" = "+val+"")
            elif col in ('Nome', 'Email'):
                query.prepare("SELECT * FROM Clientes WHERE "+col+" = '"+val+"'")
            #query.addBindValue(self.queryBox.currentText())
            #query.addBindValue(self.queryText.text())
        query.exec() 

        # Vamos popular a tabela atravessando a query linha a linha, e transpondo para a tabela célula a célula
        # Começamos o query.value no 1 para saltarmos a coluna ID (PK). Podemos também ajustar a query SQL para não buscar essa coluna
        
        self.table.clearContents()
        self.table.setRowCount(30)
        tableRow = 0
        while query.next():
            self.table.setItem(tableRow, 0, QTableWidgetItem(query.value(1)))
            self.table.setItem(tableRow, 1, QTableWidgetItem(str(query.value(2))))
            self.table.setItem(tableRow, 2, QTableWidgetItem(str(query.value(3))))
            self.table.setItem(tableRow, 3, QTableWidgetItem(str(query.value(4))))
            self.table.setItem(tableRow, 4, QTableWidgetItem(query.value(5)))
            print(query.value(0), query.value(1), query.value(2), query.value(3), query.value(4), query.value(5))

            tableRow += 1


    def add_table(self):
        table = QTableWidget(parent = self)

        table.setFixedSize(850, 500)
        table.setColumnCount(5)
        table.setSortingEnabled(False)
        table.setHorizontalHeaderLabels(['Nome','NIF', 'Telef', 'Telem', 'Email'])

        table.setColumnWidth(0, 300)
        table.setColumnWidth(1, 100)
        table.setColumnWidth(2, 100)
        table.setColumnWidth(3, 100)
        table.setColumnWidth(4, 250)
        table.move(250, 40)

    def add_search_items(self):
        queryBox = QComboBox(parent = self)
        queryBox.addItems(['Nome', 'NIF', 'Telef', 'Telem', 'Email'])
        queryBox.move(20, 60)
       
        queryText = QLineEdit(parent = self)
        queryText.move(130, 60)
        
        queryButton = QPushButton(parent = self)
        queryButton.setText('Procurar')
        queryButton.move(20, 120)

    def add_menu_bar(self):

        # Criar barra de Menu
        menuBar = self.menuBar()

        # Adicionar categorias à barra
        fileMenu = menuBar.addMenu('Ficheiro')
        helpMenu = menuBar.addMenu('Ajuda')

        # Criar Acções (botões do menu)
        ajuda1 = QAction('absolutely', self)
        ajuda2 = QAction('maidenless', self)
        ajuda3 = QAction('behaviour', self)
        ajuda4_sub = QAction("don't give up, skeleton", self)

        # Associar as acções às categorias correspondentes
        helpMenu.addAction(ajuda1)
        helpMenu.addAction(ajuda2)
        helpMenu.addAction(ajuda3)
        
        # Para adicionar sub-menus numa categoria, criamos um menu em vez de um acção
        ajuda4_menu = helpMenu.addMenu('lest...')
        ajuda4_menu.addAction(ajuda4_sub)


# Inicializar o processo do GUI
app = QApplication([])
# Inicializar main window do GUI chamando class MainWindow
mainWindow = MainWindow()
mainWindow.show()
# sys.exit termina o programa quando chamado (por exemplo quando se carrega no X) com garbage collection e memory cleanup
sys.exit(app.exec_())