#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>
#include <QDebug>
#include <QMessageBox>
#include <QCloseEvent>
#include <QStandardItemModel>
#include <QStandardItem>
#include <QFileDialog>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    /**
     * @brief mExcel указатель на эксель
     */
    QAxObject *mExcel = new QAxObject( "Excel.Application",this);
    /**
     * @brief workbooks указатель на книги экселя
     */
    QAxObject *workbooks;
    /**
     * @brief row точки
     * @brief col отсчета
     */
    int row = 1;
    int col = 1;
    /**
     * @brief numRows Количество строк
     * @brief numCols Количество столбцов
     */
    int numRows = 100;
    int numCols = 100;
    /**
     * @brief model модель для таблицы
     */
    QStandardItemModel *model = new QStandardItemModel;
    /**
     * @brief item объект для модели
     */
    QStandardItem *item = new QStandardItem(QString(""));
    /**
     * @brief filepath путь к файлу .xlsx
     */
    QString filepath;
    /**
     * @brief cellsList список ячеек
     */
    QList<QVariant> cellsList;
    /**
     * @brief rowsList список строк
     */
    QList<QVariant> rowsList;
    /**
     * @brief MainWindow конструктор главного окна
     * @param parent родительский объект
     */
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

public slots:
    /**
     * @brief setUp функция начальной установки
     */
    void setUp();
    /**
     * @brief manageXlsxFile функция подготовки файла и пути
     */
    void manageXlsxFile();
    /**
     * @brief saveToXlsx функция сохранния в файл
     */
    void saveToXlsx();
    /**
     * @brief packToXlsx функция упаковки в файл
     */
    void packToXlsx();
    /**
     * @brief getField функция создания таблицы
     */
    void getField();
    /**
     * @brief scanData фукнция изъятия данных из таблицы TableView
     */
    void scanData();
    /**
     * @brief closeEvent фукнция обработки закрытия
     * @param event событие закрытия
     */
    void closeEvent (QCloseEvent *event);
private:
    /**
     * @brief ui юай
     */
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
