#ifndef DIALOG_H
#define DIALOG_H

#include <QDialog>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QPushButton>
#include <QMenu>
#include <QMenuBar>
#include <QGroupBox>
#include <QLineEdit>
#include <QAction>
#include <QTextEdit>
#include <QLabel>
#include <QGridLayout>
#include <QFormLayout>
#include <QComboBox>
#include <QMainWindow>
#include <QList>
#include <QDebug>
#include <QTableView>
#include <QSqlQuery>
#include <QSqlTableModel>
#include <QStandardItemModel>
#include <QStandardItem>
#include <QBrush>
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
#include <QDir>
#include <QMessageBox>
#include <QAbstractItemView>
#include <QIcon>
#include <QPalette>
#include <QModelIndexList>
#include <QToolBar>
#include <QFileDialog>
using namespace QXlsx;
//#define SOLUONG 6
//#define ViTriNgay 0
//#define ViTriSo 1
//#define ViTriMaSanPham 4
//#define ViTriLoai 5
//#define ViTriMaNhanVien 3
//#define ViTriNhanVien 2

#define Tag_Ngay 1
#define Tag_So   0
#define Tag_MaNhanVien 2
#define Tag_NhanVien 3
#define Tag_MaSanPham 4
#define Tag_Loai 5

enum
{
    ViTriSo =0,
    ViTriNgay,
    ViTriNhanVien,
    ViTriMaNhanVien,
    ViTriMaSanPham,
    ViTriLoai,
    SOLUONG
};

class MainWindow : public QDialog
{
    Q_OBJECT
private:
    QMenuBar    *menuBar;
    QMenu       *fileMenu;
    QGroupBox   *gridGroup;
    QGridLayout *gridLayout;
    QLineEdit   *gridlineEdit[SOLUONG];
    QLabel      *gridLabel[SOLUONG];
    QAction     *saveAction;
    QAction     *deleteAction;
    QAction     *deleteAllAction;
    QPushButton *buttonNhap;
    QPushButton *buttonSave;
    QComboBox   *comboBox;
    QTableView  *view;
    QStandardItemModel *myModel;
    QString Ngay;
    QString So;
    QString MaSanPham;
    QString Loai;
    QString TenNhanVien;
    QString MaNhanVien;
    QFileDialog *fileDialog;
    QStandardItem *item1[SOLUONG];
    QMap<QString,QString > ThongTinNhanVien;
    QMap< int, QString > HienThiCanNhap;
    QList<QString> HienThiCombobox;
    QString fileName;

    bool firstSave;
    int SoLanNhap;
    void creatMenu();
    void creatHorizonGroupBox();
    void creategridGroupBox();
    void createFormGroupBox();
    void creatTableView();
    void readDataBase();
    void getValueFromUser();
    void createItemModel();
    void setItemModel();
    void getFileName();
private slots:
    void ButtonNhapClicked();
    void ButtonSaveClicked();
    void Delete();
    void DeleteAll();
//    void OpenFile();
public:

    MainWindow(QWidget *parent = NULL);
    ~MainWindow();
};
#endif // DIALOG_H
