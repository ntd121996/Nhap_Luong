#include "mainwindow.h"

//{
//    {ViTriNgay," Ngay: "},
//    {ViTriSo," So: "},
//    {ViTriLoai," Loai: "},
//    {ViTriNhanVien," NhanVien: "},
//    {ViTriMaSanPham," Ma SP: "},
//};
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    QVBoxLayout *mainLayout = new QVBoxLayout();
    QWidget *mainWindow = new QWidget();
    SoLanNhap = 0;
    HienThiCanNhap.insert( ViTriNgay," Ngay: ");
    HienThiCanNhap.insert( ViTriSo," So: ");
    HienThiCanNhap.insert( ViTriLoai," Loai: ");
    HienThiCanNhap.insert( ViTriNhanVien," NhanVien: ");
    HienThiCanNhap.insert( ViTriMaSanPham," Ma SP: ");
    readDataBase();
    creatMenu();
    creategridGroupBox();
    creatTableView();

    mainLayout->setMenuBar(menuBar);
    mainLayout->addWidget(this->gridGroup);
    mainLayout->addWidget(this->view);
    mainLayout->addWidget(new QLabel("Application Version 0.1. Made by DuyNT"));
    mainWindow->setLayout(mainLayout);
    this->resize(630,500);
    this->setCentralWidget(mainWindow);

    connect(buttonNhap,SIGNAL(clicked()),this,SLOT(ButtonNhapClicked()));
    connect(buttonSave,SIGNAL(clicked()),this,SLOT(ButtonSaveClicked()));
}
MainWindow::~MainWindow()
{
    delete menuBar;
    delete fileMenu;
    delete gridGroup;
    delete gridLayout;
    for ( int i = 0; i < SOLUONG; ++i )
    {
        delete gridLabel[i];
        delete gridlineEdit[i];
    }

}
void MainWindow::creatMenu()
{
    menuBar = new QMenuBar();
    fileMenu = new QMenu("File");
    exitAction = fileMenu->addAction("Exit");
    menuBar->addMenu(fileMenu);
//    connect(exitAction,&QAction::triggered,this,&QMainWindow::accept);
}
void MainWindow::creatTableView()
{
    myModel =  new  QStandardItemModel(0, 6, this);
    view = new QTableView();
    myModel->setRowCount(1);
    myModel->setColumnCount(6);

    myModel->setHeaderData(Tag_Ngay,Qt::Horizontal,tr("Ngay"),Qt::DisplayRole);
    myModel->setHeaderData(Tag_So,Qt::Horizontal,tr("So"),Qt::DisplayRole);
    myModel->setHeaderData(Tag_MaNhanVien,Qt::Horizontal,tr("Ma Nhan Vien"),Qt::DisplayRole);
    myModel->setHeaderData(Tag_NhanVien,Qt::Horizontal,tr("Nhan Vien"),Qt::DisplayRole);
    myModel->setHeaderData(Tag_MaSanPham,Qt::Horizontal,tr("Ma San Pham"),Qt::DisplayRole);
    myModel->setHeaderData(Tag_Loai,Qt::Horizontal,tr("Loai"),Qt::DisplayRole);
//    myModel->setHeaderData(Tag_Loai,Qt::Horizontal,QBrush(Qt::red),Qt::BackgroundRole);

    view->setModel(myModel);
}
void MainWindow::creategridGroupBox()
{
    gridGroup = new QGroupBox(" Nhap Luong Nhan Vien ");

    gridLayout = new QGridLayout();
    comboBox = new QComboBox();
    comboBox->addItems(ThongTinNhanVien.keys());
    buttonNhap = new QPushButton(tr(" Nhap "));
    buttonSave = new QPushButton(tr(" Luu Ra File "));
    QMap<int,QString >::iterator iterator;
    for(iterator = HienThiCanNhap.begin(); iterator != HienThiCanNhap.end(); iterator++ )
    {
        gridLabel[iterator.key()] = new QLabel(iterator.value());
        gridlineEdit[iterator.key()] = new QLineEdit();
        gridLayout->addWidget( gridLabel[ iterator.key()], iterator.key(), 0 );
        gridLayout->addWidget( gridlineEdit[ iterator.key()], iterator.key(), 1 );
    }

    gridLayout->addWidget( gridLabel[ ViTriNhanVien ], ViTriNhanVien, 0 );
    gridLayout->addWidget( comboBox, ViTriNhanVien, 1 );
    gridLayout->addWidget(buttonNhap,0,2,3,2);
    gridLayout->addWidget(buttonSave,4,2,1,2);
    gridGroup->setLayout( gridLayout );
}

void MainWindow::ButtonNhapClicked()
{
   Ngay =  gridlineEdit[ViTriNgay]->text();
   gridlineEdit[ViTriNgay]->setText(Ngay);

   So =  gridlineEdit[ViTriSo]->text();
   gridlineEdit[ViTriSo]->setText(So);

   MaSanPham =  gridlineEdit[ViTriMaSanPham]->text();
   gridlineEdit[ViTriMaSanPham]->setText(MaSanPham);

   Loai =  gridlineEdit[ViTriLoai]->text();
   gridlineEdit[ViTriLoai]->setText(Loai);


   TenNhanVien = comboBox->currentText();
   MaNhanVien = ThongTinNhanVien.value( TenNhanVien );

   item1[ViTriNgay] = new QStandardItem(Ngay);
   item1[ViTriSo] = new QStandardItem(So);
   item1[ViTriMaSanPham] = new QStandardItem(MaSanPham);
   item1[ViTriLoai] = new QStandardItem(Loai);
   item1[ViTriMaNhanVien] = new QStandardItem(MaNhanVien);
   item1[ViTriNhanVien] = new QStandardItem(TenNhanVien);

   myModel->setItem(SoLanNhap, Tag_Ngay, item1[ViTriNgay]);
   myModel->setItem(SoLanNhap, Tag_So,item1[ViTriSo]);
   myModel->setItem(SoLanNhap, Tag_MaNhanVien ,item1[ViTriMaNhanVien]);
   myModel->setItem(SoLanNhap, Tag_NhanVien,item1[ViTriNhanVien]);
   myModel->setItem(SoLanNhap, Tag_MaSanPham,item1[ViTriMaSanPham]);
   myModel->setItem(SoLanNhap, Tag_Loai,item1[ViTriLoai]);
   SoLanNhap++;
}

void MainWindow::ButtonSaveClicked()
{
    QXlsx::Document xlsxW("Test.xlsx");
    int row = 2;
    int col = 1;
    int rowCount = myModel->rowCount();
    int colCount = myModel->columnCount();
    xlsxW.write("A1", "So");
    xlsxW.write("B1", "Ngay");
    xlsxW.write("C1", "Ma NV");
    xlsxW.write("D1", "Nhan Vien");
    xlsxW.write("E1", "Ma Sp");
    xlsxW.write("F1", "Loai 1");
    for( int i = 0; i < rowCount; i++,row++)
    {
        col = 1;
        for( int j = 0; j < colCount; j++,col++)
        {
            xlsxW.write(row, col, myModel->data(myModel->index(i, j)));
        }
    }
    xlsxW.autosizeColumnWidth(1,6);
    if( xlsxW.save()) QMessageBox::information(this,tr("Export File"),tr("Luu Thanh Cong"));
//    if ( xlsxW.saveAs("Test.xlsx") ) // save the document as 'Test.xlsx'
//    {
//        qDebug() << "[debug] success to write xlsx file";
//    }
//    else
//    {
//        qDebug() << "[debug][error] failed to write xlsx file";
//    }

}

void MainWindow::readDataBase()
{
    QXlsx::Document xlsxR("ThongTin.xlsx");
    int row = 1;
    int colkey = 2;
    int colvalue = 1;
    Cell* cellkey = NULL;
    Cell* cellvalue = NULL;
    QString keyRead;
    QString valueRead;
    if ( xlsxR.load() ) // load excel file
    {
        do
        {
            cellkey = xlsxR.cellAt(row, colkey); // get cell pointer.
            cellvalue = xlsxR.cellAt(row, colvalue); // get cell pointer.
            if ( cellkey != NULL )
            {
                keyRead = cellkey->readValue().toString(); // read cell value (number(double), QDateTime, QString ...)
            }
            if ( cellvalue != NULL )
            {
                valueRead = cellvalue->readValue().toString(); // read cell value (number(double), QDateTime, QString ...)
            }
             row++;
             ThongTinNhanVien.insert( keyRead, valueRead );
        }
        while( cellkey != NULL && cellvalue != NULL );
    }
}
