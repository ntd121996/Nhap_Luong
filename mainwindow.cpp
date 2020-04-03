#include "mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QDialog(parent)
{
    QVBoxLayout *mainLayout = new QVBoxLayout();
    QMainWindow * mainWindow = new QMainWindow();
    SoLanNhap = 0;
    firstSave = true;
    fileName = "Test.xlsx";

    HienThiCanNhap.insert( ViTriNgay," Ngay: ");
//    HienThiCanNhap.insert( ViTriSo," So: ");
    HienThiCanNhap.insert( ViTriLoai," Loai: ");
    HienThiCanNhap.insert( ViTriNhanVien," NhanVien: ");
    HienThiCanNhap.insert( ViTriMaSanPham," Ma SP: ");
    readDataBase();
    creatMenu();
    creategridGroupBox();
    creatTableView();
    QToolBar *myToolBar = mainWindow->addToolBar("Tools");
    deleteAction = myToolBar->addAction("Delete");
    deleteAction->setIcon(QIcon(tr("garbage.png")));
    myToolBar->addAction(saveAction);
    mainWindow->addToolBar(myToolBar);
    mainWindow->setMenuBar(menuBar);

    mainLayout->addWidget(mainWindow);
    mainLayout->addWidget(this->gridGroup);
    mainLayout->addWidget(this->view);
    mainLayout->addWidget(new QLabel("Application Version 0.3. Made by DuyNT"));

    this->setLayout(mainLayout);
    this->resize(630,500);
    buttonNhap->setAutoDefault(true);
    buttonNhap->setDefault(false);
    buttonNhap->setPalette(QPalette(Qt::darkYellow));

    connect(buttonNhap,SIGNAL(clicked()),this,SLOT(ButtonNhapClicked()));
    connect(deleteAction,&QAction::triggered,this,&MainWindow::Delete);
    connect(deleteAllAction,&QAction::triggered,this,&MainWindow::DeleteAll);
    connect(saveAction,&QAction::triggered,this,&MainWindow::ButtonSaveClicked);
}
MainWindow::~MainWindow()
{
    delete menuBar;
    delete fileMenu;
    delete gridGroup;
    delete gridLayout;
    for ( int i = 0; i < HienThiCanNhap.size(); ++i )
    {
        delete gridLabel[i];
    }
    delete gridlineEdit[ViTriNgay];
    delete gridlineEdit[ViTriMaSanPham];
    delete gridlineEdit[ViTriSo];
    delete gridlineEdit[ViTriLoai];
    delete buttonNhap;

    delete comboBox;
    delete item1[ViTriNgay];
    delete item1[ViTriSo];
    delete item1[ViTriMaSanPham] ;
    delete item1[ViTriLoai];
    delete item1[ViTriMaNhanVien];
    delete item1[ViTriNhanVien];
    delete myModel;
    delete view;
    delete fileDialog;
}
void MainWindow::creatMenu()
{
    menuBar = new QMenuBar();
    fileMenu = new QMenu("File");
    saveAction = fileMenu->addAction("Save");
    deleteAllAction = fileMenu->addAction("Delete All");
    deleteAllAction->setIcon(QIcon(tr("delete.png")));
    saveAction->setIcon(QIcon(tr("save.png")));
    menuBar->addMenu(fileMenu);
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

    comboBox->addItems(HienThiCombobox);
    comboBox->view()->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOn);
    buttonNhap = new QPushButton(tr(" Nhap "));
    QMap<int,QString >::iterator iterator;
    for(iterator = HienThiCanNhap.begin(); iterator != HienThiCanNhap.end(); iterator++ )
    {
        gridLabel[iterator.key()] = new QLabel(iterator.value());
        if(iterator.key() == ViTriNhanVien) continue;
        gridlineEdit[iterator.key()] = new QLineEdit();
        gridLayout->addWidget( gridLabel[ iterator.key()], iterator.key(), 0 );
        gridLayout->addWidget( gridlineEdit[ iterator.key()], iterator.key(), 1);
    }

    gridLayout->addWidget( gridLabel[ ViTriNhanVien ], ViTriNhanVien, 0 );
    gridLayout->addWidget( comboBox, ViTriNhanVien, 1 );
    gridLayout->addWidget(buttonNhap,0,2,3,1);
    gridGroup->setLayout( gridLayout );
}

void MainWindow::ButtonNhapClicked()
{

   QStringList splitString;
   QList<QString> Sum;
   int tong = 0;
   Ngay = gridlineEdit[ViTriNgay]->text();
   if(!Ngay.isEmpty())
   {
       So = 'N'+Ngay.left(2);
   }
   qDebug() << "So:" << So;
   MaSanPham = gridlineEdit[ViTriMaSanPham]->text();

   Loai = gridlineEdit[ViTriLoai]->text();
   gridlineEdit[ViTriLoai]->clear();
   splitString = Loai.split(' ',QString::SkipEmptyParts);
   while( !splitString.isEmpty())
   {
       Sum.append(splitString.takeFirst());
       tong += Sum.takeFirst().toInt();
   }
   qDebug() << "Tong:" << tong;
   if(tong != 0 )Loai = QString::number(tong);

   TenNhanVien = comboBox->currentText();
   MaNhanVien = ThongTinNhanVien.value( TenNhanVien );

   item1[ViTriNgay] = new QStandardItem(Ngay);
   item1[ViTriSo] = new QStandardItem(So);
   item1[ViTriMaSanPham] = new QStandardItem(MaSanPham);
   item1[ViTriLoai] = new QStandardItem();
   item1[ViTriMaNhanVien] = new QStandardItem(MaNhanVien);
   item1[ViTriNhanVien] = new QStandardItem(TenNhanVien);

   myModel->setItem(SoLanNhap, Tag_Ngay, item1[ViTriNgay]);
   myModel->setItem(SoLanNhap, Tag_So,item1[ViTriSo]);
   myModel->setItem(SoLanNhap, Tag_MaNhanVien ,item1[ViTriMaNhanVien]);
   myModel->setItem(SoLanNhap, Tag_NhanVien,item1[ViTriNhanVien]);
   myModel->setItem(SoLanNhap, Tag_MaSanPham,item1[ViTriMaSanPham]);
   myModel->setItem(SoLanNhap, Tag_Loai,item1[ViTriLoai]);
   if(tong != 0 )myModel->setData(myModel->index(SoLanNhap, ViTriLoai ),QVariant(tong));

   view->scrollToBottom();
   SoLanNhap++;

}

void MainWindow::ButtonSaveClicked()
{

    QStringList splitString;
    QString Temp;
    if(firstSave)
    {
        fileDialog = new QFileDialog();
        fileDialog->exec();
        Temp = fileDialog->selectedFiles().takeFirst();
        splitString = Temp.split('/',QString::SkipEmptyParts);
        Temp = splitString.takeLast();
        firstSave = false;
    }
    if( Temp.compare(Temp.right(4),"xlsx") == 0 )
    {
        fileName = Temp;
    }
    QXlsx::Document xlsxW(fileName);
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
    xlsxW.selectSheet("Sheet1");
    for( int i = 0; i < rowCount; i++,row++)
    {
        col = 1;
        for( int j = 0; j < colCount; j++,col++)
        {
            xlsxW.write(row, col, myModel->data(myModel->index(i, j)));
        }
    }
    xlsxW.autosizeColumnWidth(1,6);
    if( xlsxW.save()) QMessageBox::information(this,tr("Luu Thanh Cong"),tr("Have A Nice Day My Darling <3 !!! "));

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
            if( cellkey == NULL || cellvalue == NULL )
            {
                break;
            }
            keyRead = cellkey->readValue().toString(); // read cell value (number(double), QDateTime, QString ...)
            valueRead = cellvalue->readValue().toString(); // read cell value (number(double), QDateTime, QString ...)
            row++;
            HienThiCombobox.append(keyRead);
            ThongTinNhanVien.insert( keyRead, valueRead );
        }
        while( cellkey != NULL && cellvalue != NULL );
    }
}

void MainWindow::Delete()
{
    QModelIndexList selected = view->selectionModel()->selectedRows();
    int  count = selected.count();
    int i = 0;
    for( i = count ; i > 0;i --)
    {
        myModel->removeRow(selected.at(i - 1).row());
    }
    SoLanNhap -= count;
}
void MainWindow::DeleteAll()
{
    int rowCount = myModel->rowCount();
    int colCount = myModel->columnCount();
    qDebug() <<"Row:" <<rowCount << "Col:" <<colCount;
    for( int i = 0; i < rowCount; i++)
    {

        for( int j = 0; j < colCount; j++)
        {
            myModel->takeItem(i, j);
        }
    }

    myModel->setRowCount(0);
    myModel->setColumnCount(6);
    SoLanNhap = 0;
}
