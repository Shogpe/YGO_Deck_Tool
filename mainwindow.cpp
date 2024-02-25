#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QStandardItemModel>
#include <QDebug>
#include <QAxObject>
#include <QFileDialog>
#include <QDesktopServices>
#include <qtablewidget.h>

# pragma execution_character_set("utf-8")

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    for(int idx = 0; idx<CORE_GROUP_MAX; idx++)
    {
        coreCard[idx].num[0] = 0;
        coreCard[idx].num[1] = 0;
        coreCard[idx].num[2] = 0;
        coreCard[idx].kinds = 0;
        coreCard[idx].valid = 0;
    }
    decksNum = 0;
    sKengNum = 0;
    hKengNum = 0;
    otherNum = 0;

    for(int i = 0; i<DRAW_CASE;i++)
    {
        sKengExpect[i] = 0;
        hKengExpect[i] = 0;
        SingExpect[i] = 0;
        DoubExpect[i] = 0;
        TribExpect[i] = 0;
        KegRateExpan0[i] = 0;
        KegRateExpan1[i] = 0;
        KegRateExpan2[i] = 0;
        uselessNum[i] = 0;
    }

    // 标题菜单栏
    connect(ui->menubar, &QMenuBar::triggered, this, [=](QAction* action) {
        qDebug() << action->objectName();
        this->deal_msg(action->objectName());
    });

    ui->inPutWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);    //填充表格
    ui->inPutWidget->setEditTriggers(QAbstractItemView::CurrentChanged);

    ui->outPutWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);    //填充表格
    ui->outPutWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);


    runTimer.setInterval(2000);
    connect(&runTimer, &QTimer::timeout, this, &MainWindow::RunSys);
    runTimer.start();
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::deal_msg(const QString& msg)
{
    if(msg == "actionimport")
    {
        ImportExcel();
    }
    if(msg == "actionexport")
    {

        ExportExcel();
    }
}

void MainWindow::InputCheck()
{
    uint8_t  cardNum = 0;

    decksNum = ui->deckNum->value();
    sKengNum = ui->shoukengNum->value();
    hKengNum = ui->houkengNum->value();
    cardNum += sKengNum + hKengNum;

    otherNum = 0;

    for(int col = 0; col<ui->inPutWidget->columnCount();col++)
    {
        if(ui->inPutWidget->item(0,col) != NULL)
        {
            coreCard[col].kinds = ui->inPutWidget->item(0,col)->text().toUInt();

            if(coreCard[col].kinds>3)
            {
                coreCard[col].kinds = 0;
                coreCard[col].valid = 0;
            }
        }
        else
        {
            coreCard[col].kinds = 0;
            coreCard[col].valid = 0;
        }

        uint8_t kindCnt = 0;
        for(int i = 1; i <= 3; i++)
        {
            if(kindCnt>=coreCard[col].kinds)
            {
                break;
            }

            if(ui->inPutWidget->item(i,col)!=NULL)
            {
                uint8_t thisCardNum = ui->inPutWidget->item(i,col)->text().toUInt();
                if(thisCardNum != 0)
                {
                    coreCard[col].num[kindCnt++] = thisCardNum;
                }
            }
        }

        if(kindCnt == coreCard[col].kinds && coreCard[col].kinds != 0)
        {
            coreCard[col].valid = 1;
        }
        else
        {
            coreCard[col].valid = 0;
        }
    }

    for(int idx = 0; idx<CORE_GROUP_MAX; idx++)
    {
        if(coreCard[idx].valid)
        {
            cardNum += coreCard[idx].num[0];
            cardNum += coreCard[idx].num[1];
            cardNum += coreCard[idx].num[2];
            qDebug()<<idx<<" : "<<coreCard[idx].kinds<<"|"<<coreCard[idx].num[0]<<"-"<<coreCard[idx].num[1]<<"-"<<coreCard[idx].num[2];
        }
    }

    otherNum = decksNum - cardNum;

    QString text;
    if(otherNum>=0)
    {
        text = QString("未统计的\n卡牌数：\n");
        text = text+QString::number(otherNum);
        ui->NumLable->setText(text);
        ui->NumLable->setStyleSheet("color:green");
    }
    else
    {
        ui->NumLable->setText("卡牌数量\n不正确\n结果不可用");
        ui->NumLable->setStyleSheet("color:red");
    }
}


static unsigned long long int C(uint8_t m,uint8_t n)
{
    //C(7,20)==77520
    unsigned long long int res = 1;
    if(m==0||n==0||m==n)
    {
        return 1;
    }

    if(m>n)
    {
        return 0;
    }

    if(m>n/2)
    {
        m = n-m;
    }

    //n>m
    uint8_t j = 1;

    for(uint8_t i =m+1; i<=n; i++)
    {
        res*= i;
        if(res%j == 0&&j<=(n-m))
        {
            res/=j;
            j++;
        }

        if(res%j == 0&&j<=(n-m))
        {
            res/=j;
            j++;
        }

        if(res%j == 0&&j<=(n-m))
        {
            res/=j;
            j++;
        }
    }

    //继续没除完的
    for(; j<=(n-m); j++)
    {
        res/=j;
    }

    if(res == 0)
    {
        res = 1;
    }
    return res;
}

static double CalExpectSimple( uint8_t num , uint8_t sel, uint8_t tol)
{
    double res = 0;
    double p = 0;
    for(int x = 0; x <= sel; x++)
    {
        if(x>num||sel-x>tol-num)
        {
            p = 0;
        }
        else
        {
            p = double(C(x,num))*double(C(sel-x,tol-num))/C(sel,tol);
        }
        res += p*x;
    }
    return res;
}

static double CalExpectSingle( uint8_t num , uint8_t sel, uint8_t tol, double* pPCache)
{
    double res = 0;
    double p = 0;
    for(int x = 0; x <= sel; x++)
    {
        if(x>num||sel-x>tol-num)
        {
            p = 0;
        }
        else
        {
            p = double(C(x,num)*C(sel-x,tol-num))/C(sel,tol);
        }
        pPCache[x] = p;
        res += p*x;
    }
    return res;
}

static double CalExpectDouble( uint8_t num1 , uint8_t num2 , uint8_t sel, uint8_t tol, double* pPCache)
{
    double res = 0;
    double p = 0;
    double pLast = 0;
    uint8_t min = num1<=num2?num1:num2;
    uint8_t max = num1>=num2?num1:num2;
    uint8_t other = tol-num1-num2;

    qDebug()<<num1<<"=="<<num2<<"=="<<sel<<"=="<<tol;
    for(int x = 0; x < DEFAULT_HANDS+DRAW_CASE; x++)
    {
        pPCache[x] = 0;
    }

    for(int x = sel/2; x >= 0; x--)
    {
        p = 0;
        if(x>min||2*x>sel)
        {
            p = 0;
        }
        else
        {
              for(uint8_t i1 = x; i1 <= num1; i1++)
              {
                  for(uint8_t i2 = x; i2 <= num2; i2++)
                  {
                      if(i1+i2<=sel)
                      {
                          p  += double(C(i1,num1)*C(i2,num2)*C(sel-i1-i2,other))/C(sel,tol);
                      }
                  }
              }
        }
        pPCache[x] = p - pLast;
        pLast = p;
        p = pPCache[x];


        qDebug()<<"p"<<x<<" = "<<p;
        res += p*x;
    }
    return res;
}

static double CalExpectTriple( uint8_t num1 , uint8_t num2 , uint8_t num3 ,uint8_t sel, uint8_t tol, double* pPCache)
{
    double res = 0;
    double p = 0;
    uint8_t min = num1<=num2?num1:num2;
    min = min<=num3?min:num3;

    qDebug()<<num1<<"=="<<num2<<"=="<<num3<<"=="<<sel<<"=="<<tol;
    for(int x = 0; x <= sel; x++)
    {
        if(x>min||3*x>sel)
        {
            p = 0;
        }
        else
        {
            p = double(C(x,num1)*C(x,num2)*C(x,num3)*C(sel-3*x,tol-3*x))/C(num1-x,num1)/C(num2-x,num2)/C(num3-x,num3)/C(sel,tol);
        }
        pPCache[x] = p;

        qDebug()<<"p"<<x<<" = "<<p;
        res += p*x;
    }
    return res;
}


void MainWindow::CalRate()
{
    if(otherNum>=0)
    {
        for(int i = 0; i<DRAW_CASE;i++)
        {
            uint8_t handCard = DEFAULT_HANDS + i;

            sKengExpect[i] = CalExpectSimple(sKengNum, handCard, decksNum);
            hKengExpect[i] = CalExpectSimple(hKengNum, handCard, decksNum);

            SingExpect[i] = 0;
            for(int group = 0; group < CORE_GROUP_MAX; group++)
            {
                if(coreCard[group].valid == 1&&coreCard[group].kinds ==1)
                {
                    SingExpect[i] += CalExpectSingle(coreCard[group].num[0], handCard, decksNum, SingleP[group]);
                }
            }

            DoubExpect[i] = 0;
            for(int group = 0; group < CORE_GROUP_MAX; group++)
            {
                if(coreCard[group].valid == 1&&coreCard[group].kinds ==2)
                {
                    DoubExpect[i] += CalExpectDouble(coreCard[group].num[0],coreCard[group].num[1], handCard, decksNum, DoubleP[group]);
                }
            }

            TribExpect[i] = 0;
            for(int group = 0; group < CORE_GROUP_MAX; group++)
            {
                if(coreCard[group].valid == 1&&coreCard[group].kinds ==3)
                {
                    TribExpect[i] += CalExpectTriple(coreCard[group].num[0],coreCard[group].num[1],coreCard[group].num[2], handCard, decksNum, TribleP[group]);
                }
            }

//            KegRateExpan0[i] = 0;
//            for(uint8_t j = 1; j <= handCard; j++)
//            {
//                for(uint8_t k = 1; k < 10; k++)
//                {
//                    KegRateExpan0[i] += pSingle[k][j];
//                }
////                uint8_t num1,num2,num3 =0;
////                for(num1 = 0;num1 <= j;num1++)
////                {
////                    for(num2 = 0;num2 <= j;num2++)
////                    {
////                        for(num3 = 0;num3 <= j;num3++)
////                        {
////                            if(num1+num2+num3 == j)
////                            {
////                                KegRateExpan0[i] += pSingle[num1] * pDouble[num2] * pTrible[num1];
////                            }
////                        }
////                    }
////                }
//            }
//            if(KegRateExpan0[i]>1)
//            {
//                KegRateExpan0[i] = 1;
//            }

//            KegRateExpan1[i] = 0;
//            for(uint8_t j = 2; j < handCard; j++)
//            {
//                for(uint8_t k = 1; k < 10; k++)
//                {
//                    KegRateExpan0[i] += pSingle[k][j];
//                }
////                uint8_t num1,num2,num3 =0;
////                for(num1 = 0;num1 <= j;num1++)
////                {
////                    for(num2 = 0;num2 <= j;num2++)
////                    {
////                        for(num3 = 0;num3 <= j;num3++)
////                        {
////                            if(num1+num2+num3 == j)
////                            {
////                                KegRateExpan0[i] += pSingle[num1] * pDouble[num2] * pTrible[num1];
////                            }
////                        }
////                    }
////                }
//            }
//            if(KegRateExpan1[i]>1)
//            {
//                KegRateExpan1[i] = 1;
//            }

//            KegRateExpan2[i] = 0;
//            for(uint8_t j = 3; j < handCard; j++)
//            {
//                for(uint8_t k = 1; k < 10; k++)
//                {
//                    KegRateExpan0[i] += pSingle[k][j];
//                }
////                uint8_t num1,num2,num3 =0;
////                for(num1 = 0;num1 <= j;num1++)
////                {
////                    for(num2 = 0;num2 <= j;num2++)
////                    {
////                        for(num3 = 0;num3 <= j;num3++)
////                        {
////                            if(num1+num2+num3 == j)
////                            {
////                                KegRateExpan0[i] += pSingle[num1] * pDouble[num2] * pTrible[num1];
////                            }
////                        }
////                    }
////                }

//            }
//            if(KegRateExpan2[i]>1)
//            {
//                KegRateExpan2[i] = 1;
//            }



            uselessNum[i] = handCard - sKengExpect[i] - hKengExpect[i] - SingExpect[i] - 2*DoubExpect[i] - 3*TribExpect[i];
            if(uselessNum[i]<0)
            {
                uselessNum[i] = 0;
            }

        }

//        qDebug()<<"------";
//        qDebug()<<C(7,120);
//        qDebug()<<"------";
    }
    else
    {
        for(int i = 0; i<DRAW_CASE;i++)
        {
            sKengExpect[i] = 0;
            hKengExpect[i] = 0;
            SingExpect[i] = 0;
            DoubExpect[i] = 0;
            TribExpect[i] = 0;
            KegRateExpan0[i] = 0;
            KegRateExpan1[i] = 0;
            KegRateExpan2[i] = 0;
            uselessNum[i] = 0;
        }
    }
}

void MainWindow::RefreshDisplay()
{
    for(int col = 0; col<ui->outPutWidget->columnCount();col++)
    {
        ui->outPutWidget->setItem(0,col,new QTableWidgetItem(QString::number(sKengExpect[col])));
        ui->outPutWidget->setItem(1,col,new QTableWidgetItem(QString::number(hKengExpect[col])));
        ui->outPutWidget->setItem(2,col,new QTableWidgetItem(QString::number(SingExpect[col])));
        ui->outPutWidget->setItem(3,col,new QTableWidgetItem(QString::number(DoubExpect[col])));
        ui->outPutWidget->setItem(4,col,new QTableWidgetItem(QString::number(TribExpect[col])));
        ui->outPutWidget->setItem(5,col,new QTableWidgetItem(QString::number(KegRateExpan0[col])));
        ui->outPutWidget->setItem(6,col,new QTableWidgetItem(QString::number(KegRateExpan1[col])));
        ui->outPutWidget->setItem(7,col,new QTableWidgetItem(QString::number(KegRateExpan2[col])));
        ui->outPutWidget->setItem(8,col,new QTableWidgetItem(QString::number(uselessNum[col])));
    }
}

void MainWindow::RunSys()
{
    InputCheck();
    CalRate();
    RefreshDisplay();
}


void MainWindow::ImportExcel()
{
    QString strFile = QFileDialog::getOpenFileName(this,QStringLiteral("save"),"",tr("Exel file(*.xls *.xlsx)"));
    qDebug()<<strFile;
    if (strFile.isEmpty())
    {
        return;
    }

    QAxObject excel("Excel.Application"); //加载Excel驱动
    excel.setProperty("Visible", false);//不显示Excel界面，如果为true会看到启动的Excel界面
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)", strFile); //打开指定文件
    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //获取工作表
    QString ExcelName;
    static int row_count = 0,column_count = 0;
    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目,如下图，有 3 页


    if(sheet_count > 0)
    {
        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1); //设置为 获取第一页 数据
        QAxObject *used_range = work_sheet->querySubObject("UsedRange");
        QAxObject *rows = used_range->querySubObject("Rows");
        row_count = rows->property("Count").toInt();  //获取行数

        QAxObject *column = used_range->querySubObject("Columns");
        column_count = column->property("Count").toInt();  //获取列数
        qDebug() << row_count<<"-"<< column_count;
        //获取第一行第一列数据
        ExcelName = work_sheet->querySubObject("Cells(int,int)", 1,1)->property("Value").toString();
        qDebug() << ExcelName;
        //获取表格中需要的数据
        ui->inPutWidget->setRowCount(row_count - 1);
        for (int i =2; i <= row_count; i++)
        {
            for (int j = 1; j <= column_count;j++)
            {
                qDebug() << "reading";
                //QString txt = work_sheet->querySubObject("Cells(int,int)", i,j)->property("Value2()").toString(); //获取单元格内容
                QString txt = work_sheet->querySubObject("Cells(int,int)", i,j)->dynamicCall("Value2()").toString();

                QTableWidgetItem* p_item = new QTableWidgetItem(txt);

                ui->inPutWidget->setItem(i-2,j-1,p_item );
                qDebug() << "row "<<i-2<<",col "<<j-1<<" set to :"<<txt;


            }
        }
        ui->inPutWidget->repaint();

        work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
        excel.dynamicCall("Quit(void)");  //退出
    }



}

void MainWindow::ExportExcel()
{
    //获取保存路径
       QString filepath=QFileDialog::getSaveFileName(this,tr("Save"),".",tr(" (*.xlsx)"));
       if(!filepath.isEmpty()){
           QAxObject *excel = new QAxObject(this);
           //连接Excel控件
           excel->setControl("Excel.Application");
           //不显示窗体
           excel->dynamicCall("SetVisible (bool Visible)","false");
           //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
           excel->setProperty("DisplayAlerts", false);
           //获取工作簿集合
           QAxObject *workbooks = excel->querySubObject("WorkBooks");
           //新建一个工作簿
           workbooks->dynamicCall("Add");
           //获取当前工作簿
           QAxObject *workbook = excel->querySubObject("ActiveWorkBook");
           //获取工作表集合
           QAxObject *worksheets = workbook->querySubObject("Sheets");
           //获取工作表集合的工作表1，即sheet1
           QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);
           //设置表头值
           for(int i=1;i<ui->inPutWidget->columnCount()+1;i++)
           {
               //设置设置某行某列
               QAxObject *Range = worksheet->querySubObject("Cells(int,int)", 1, i);
               if(ui->inPutWidget->horizontalHeaderItem(i-1) == NULL)
               {
                   Range->dynamicCall("SetValue(const QString &)",QString("组")+QString::number(i));
               }
               else
               {
                   Range->dynamicCall("SetValue(const QString &)",ui->inPutWidget->horizontalHeaderItem(i-1)->text());
               }

           }
           //设置表格数据
           for(int i = 1;i<ui->inPutWidget->rowCount()+1;i++)
           {
               for(int j = 1;j<ui->inPutWidget->columnCount()+1;j++)
               {
                   QAxObject *Range = worksheet->querySubObject("Cells(int,int)", i+1, j);
                   if(ui->inPutWidget->item(i-1,j-1) == NULL)
                   {
                       //qDebug() << i <<"-"<< j <<"- NULL";
                       Range->dynamicCall("SetValue(const QString &)",QString("0"));
                   }
                   else
                   {
                       //qDebug() << i <<"-"<< j <<ui->inPutWidget->item(i-1,j-1)->data(Qt::DisplayRole).toString();
                       Range->dynamicCall("SetValue(const QString &)",ui->inPutWidget->item(i-1,j-1)->data(Qt::DisplayRole).toString());
                   }
               }
           }
           workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));//保存至filepath
           workbook->dynamicCall("Close()");//关闭工作簿
           excel->dynamicCall("Quit()");//关闭excel
           delete excel;
           excel=NULL;
           qDebug() << "\n导出成功！！！";
       }

}
