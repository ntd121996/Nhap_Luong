#include "mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QDialog(parent)
{
    QVBoxLayout *mainLayout = new QVBoxLayout();
    QMainWindow * mainWindow = new QMainWindow();
    SoLanNhap = 0;
    firstSave = true;
    fileName.clear();
//    qDebug()/* << 0 + rand()% (9 + 1 - 0);*/
    HienThiCanNhap.insert( ViTriNgay," Ngay: ");
//    HienThiCanNhap.insert( ViTriSo," So: ");
    HienThiCanNhap.insert( ViTriLoai," Loai: ");
    HienThiCanNhap.insert( ViTriNhanVien," NhanVien: ");
    HienThiCanNhap.insert( ViTriMaSanPham," Ma SP: ");
    LoiChuc.append("Have A Nice Day My Dear <3 ");
    LoiChuc.append("Nguyen Thai Duy Love Truong My Hoa ");
    LoiChuc.append("Luon Vui Ve Nhe Em Yeu <3 ");
    LoiChuc.append("Anh Nho Em, Vo Iu A ,Hiu Hiu T.T !!!! ");
    LoiChuc.append("Wo Ai Ni, Saranghaeyo, Tiamo <3 ");
    LoiChuc.append("I've Died Every Day Because Miss U T.T ");
    LoiChuc.append("Keep Calm And Love Thai Duy Much More, My Love <3 ");
    LoiChuc.append("A Khong Luon Leo Dau Nhe, Chi Yeu Em Thoi,Hi Hi :p");
    LoiChuc.append("Em Cuoi Dep Gap Van Lan Luc Em Cau Gian, Be Smile ^_^ ");
    LoiChuc.append("I Will Love U Every Morning !!!");

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
    mainLayout->addWidget(new QLabel("Application Version 0.. Made by DuyNT"));

    this->setLayout(mainLayout);
    this->resize(630,500);
    buttonNhap->setAutoDefault(true);
    buttonNhap->setDefault(false);
    buttonNhap->setPalette(QPalette(Qt::darkYellow));
    buttonNhap->setFocusPolicy(Qt::ClickFocus);
    this->gridlineEdit[ViTriMaSanPham]->setFocus();
    this->gridlineEdit[ViTriNgay]->setFocusPolicy(Qt::ClickFocus);
    this->view->setFocusPolicy(Qt::NoFocus);


    connect(buttonNhap,SIGNAL(clicked()),this,SLOT(ButtonNhapClicked()));
    connect(deleteAction,&QAction::triggered,this,&MainWindow::Delete);
    connect(deleteAllAction,&QAction::triggered,this,&MainWindow::DeleteAll);
    connect(saveAction,&QAction::triggered,this,&MainWindow::ButtonSaveClicked);
}
MainWindow::~MainWindow()
{
    delete menuBar;
    delete fileMenu;
    delete gridLayout;
    delete buttonNhap;
    delete comboBox;
    delete gridGroup;
    delete myModel;
    delete view;
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

    gridLayout->addWidget(buttonNhap,0,2,3,1);
    gridLayout->addWidget( gridLabel[ ViTriNhanVien ], ViTriNhanVien, 0 );
    gridLayout->addWidget( comboBox, ViTriNhanVien, 1 );

    gridGroup->setLayout( gridLayout );
}
void MainWindow::ButtonNhapClicked()
{
   getValueFromUser();
   createItemModel();
   setItemModel();
   view->scrollToBottom();
   this->comboBox->setFocus();


   SoLanNhap++;
}
void MainWindow::ButtonSaveClicked()
{
    getFileName();
    if( fileName.isNull())
    {
        fileName = "Test.xlsx";
    }
    QXlsx::Document xlsxW(fileName);
    int rowCount = myModel->rowCount();
    int colCount = myModel->columnCount();
    int rowIndex = 0;
    int colIndex = 0;
    int rowWrite = 0;
    int colWrite = 0;
    if( fileName.compare(fileName.right(9),"Test.xlsx") == 0 )
    {
        rowWrite = 2;
        colWrite = 1;
        xlsxW.write("A1", "So");
        xlsxW.write("B1", "Ngay");
        xlsxW.write("C1", "Ma NV");
        xlsxW.write("D1", "Nhan Vien");
        xlsxW.write("E1", "Ma Sp");
        xlsxW.write("F1", "Loai 1");
        xlsxW.selectSheet("Sheet1");
        rowIndex = rowWrite;
        colIndex = colWrite;
        for( int i = 0; i < rowCount; i++,rowIndex++)
        {
            colIndex = colWrite;
            for( int j = 0; j < colCount; j++,colIndex++)
            {
                xlsxW.write(rowIndex, colIndex, myModel->data(myModel->index(i, j)));
            }
        }
        xlsxW.autosizeColumnWidth(1,6);
    }
    // File is exist
    else
    {
        xlsxW.load();
        xlsxW.selectSheet("Nhập");
        Cell* cellkey = NULL;
        int nextRow = 6;
        do
        {
            cellkey = xlsxW.cellAt(nextRow, 4); // get cell pointer.
            if( cellkey == NULL )
            {
                break;
            }
            if( cellkey->readValue().isNull())
            {
                break;
            }
            nextRow++;
        }
        while(cellkey != NULL);
//        qDebug() << row;
//        Select Sheet to write

//        Set up row,column to write
        rowWrite = nextRow;
        colWrite = 1;

        rowIndex = rowWrite;
        colIndex = colWrite;
        for( int i = 0; i < rowCount; i++,rowIndex++)
        {
            colIndex = colWrite;
            for( int j = 0; j < colCount; j++,colIndex++)
            {
                if(colIndex == 6 ) colIndex = 8;
                xlsxW.write(rowIndex, colIndex, myModel->data(myModel->index(i, j)));
            }
        }
    }

    if( xlsxW.save()) QMessageBox::information(this,tr("Luu Thanh Cong"),LoiChuc.at((0 + rand()% (9 + 1 - 0))));

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
//    qDebug() <<"Row:" <<rowCount << "Col:" <<colCount;
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
void MainWindow::getValueFromUser()
{
    QStringList splitString;
    QList<QString> Sum;
    int tong = 0;
    Ngay = gridlineEdit[ViTriNgay]->text();
    if(!Ngay.isEmpty())
    {
        So = 'N'+Ngay.left(2);
    }
    MaSanPham = gridlineEdit[ViTriMaSanPham]->text();
    qDebug() << MaSanPham;
    Loai = gridlineEdit[ViTriLoai]->text();
    gridlineEdit[ViTriLoai]->clear();
    splitString = Loai.split(' ',QString::SkipEmptyParts);
    while( !splitString.isEmpty())
    {
        Sum.append(splitString.takeFirst());
        tong += Sum.takeFirst().toInt();
    }
    if(tong != 0) Loai = QString::number(tong);
    TenNhanVien = comboBox->currentText();
    MaNhanVien = ThongTinNhanVien.value( TenNhanVien );
}
void MainWindow::createItemModel()
{
    item1[ViTriNgay] = new QStandardItem(Ngay);
    item1[ViTriSo] = new QStandardItem(So);
    item1[ViTriMaSanPham] = new QStandardItem(MaSanPham);

    item1[ViTriLoai] = new QStandardItem(Loai);
    item1[ViTriMaNhanVien] = new QStandardItem(MaNhanVien);
    item1[ViTriNhanVien] = new QStandardItem(TenNhanVien);
//    qDebug() << MaSanPham << item1[ViTriMaSanPham];
}
void MainWindow::setItemModel()
{
    myModel->setItem(SoLanNhap, Tag_Ngay, item1[ViTriNgay]);
    myModel->setItem(SoLanNhap, Tag_So,item1[ViTriSo]);
    myModel->setItem(SoLanNhap, Tag_MaNhanVien ,item1[ViTriMaNhanVien]);
    myModel->setItem(SoLanNhap, Tag_NhanVien,item1[ViTriNhanVien]);
    myModel->setItem(SoLanNhap, Tag_MaSanPham,item1[ViTriMaSanPham]);
    myModel->setItem(SoLanNhap, Tag_Loai,item1[ViTriLoai]);
    if(Loai.toInt()!= 0)
    {
        myModel->setData(myModel->index(SoLanNhap, Tag_Loai ),QVariant(Loai.toInt()));
    }

}
void MainWindow::getFileName()
{
    QStringList splitString;
    QString Temp;
    fileDialog = new QFileDialog(this);
    fileDialog->setNameFilter(tr("Filter File (*.png *.xlsx)"));
    fileDialog->setViewMode(QFileDialog::Detail);
    fileDialog->exec();
    fileName = fileDialog->selectedFiles().takeFirst();
    Temp = fileName;
    splitString = Temp.split('/',QString::SkipEmptyParts);
    Temp = splitString.takeLast();
    if( Temp.compare(Temp.right(4),"xlsx") != 0 )
    {
        fileName.clear();
    }
}
