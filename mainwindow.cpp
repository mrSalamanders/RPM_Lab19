#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    this->setUp();
}

void MainWindow::setUp()
{
    ui->setupUi(this);
    this->ui->tableView->setModel(this->model);
    this->mExcel = new QAxObject( "Excel.Application",this);
    this->workbooks = mExcel->querySubObject( "Workbooks" );
    connect(this->ui->pushButton, &QPushButton::clicked, this, &MainWindow::saveToXlsx);
    this->getField();
}

void MainWindow::manageXlsxFile()
{
    this->filepath = QFileDialog::getSaveFileName(this, "Save File", "", ".xlsx(*.xlsx)");
    qDebug() << filepath;
}

void MainWindow::saveToXlsx()
{
    manageXlsxFile();
    scanData();
    packToXlsx();
}

void MainWindow::packToXlsx()
{
    QAxObject *workbook = this->workbooks->querySubObject("Add()");
    workbook->querySubObject("SaveAs(const QString&)", QDir::toNativeSeparators(this->filepath));
    QAxObject *mSheets = workbook->querySubObject( "Sheets" );
    QAxObject *StatSheet = mSheets->querySubObject( "Item(const QVariant&)", QVariant("Лист1") );
    QAxObject* Cell1 = StatSheet->querySubObject("Cells(QVariant&,QVariant&)", row, col);
    QAxObject* Cell2 = StatSheet->querySubObject("Cells(QVariant&,QVariant&)", row + numRows - 1, col + numCols - 1);
    QAxObject* range = StatSheet->querySubObject("Range(const QVariant&,const QVariant&)", Cell1->asVariant(), Cell2->asVariant() );

    range->setProperty("Value", QVariant(rowsList) );

    delete range;
    delete Cell1;
    delete Cell2;

    delete StatSheet;
    delete mSheets;
    workbook->dynamicCall("Close(Boolean)", true);
    delete workbook;
}

void MainWindow::getField()
{
    for(int i = 0; i < this->numRows; i++) {
        for(int j = 0; j < this->numCols; j++) {
            model->setItem(i, j, item);
        }
    }
}

void MainWindow::scanData()
{
    this->rowsList.clear();
    for(int i = 0; i < this->numRows; i++) {
        cellsList.clear();
        for(int j = 0; j < this->numCols; j++) {
            this->cellsList << QVariant(ui->tableView->model()->data(ui->tableView->model()->index(i, j), Qt::DisplayRole));
        }
        this->rowsList << QVariant(cellsList);
    }
}

void MainWindow::closeEvent (QCloseEvent *event)
{
    QMessageBox::StandardButton resBtn = QMessageBox::question( this, "A?",
                                                                tr("Want quit?\n"),
                                                                QMessageBox::No | QMessageBox::Yes,
                                                                QMessageBox::Yes);
    if (resBtn != QMessageBox::Yes) {
        event->ignore();
    } else {
        this->mExcel->dynamicCall("Quit()");
        delete this->workbooks;
        delete this->mExcel;
        event->accept();
    }
}

MainWindow::~MainWindow()
{
    delete ui;
}

