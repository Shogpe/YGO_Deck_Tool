#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTimer>

#define DECK_NUM_MIN    40
#define DECK_NUM_MAX    60

#define DEFAULT_HANDS   5

#define DRAW_0          0
#define DRAW_1          1
#define DRAW_2          2

#define DRAW_CASE       3

#define CORE_GROUP_MAX  10

typedef struct
{
    uint8_t valid;
    uint8_t kinds;
    uint8_t num[3];

}coreCard_t;



QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void deal_msg(const QString& msg);

    void InputCheck();
    void CalRate();
    void RefreshDisplay();

    void RunSys();

    void ImportExcel();
    void ExportExcel();

private:
    Ui::MainWindow *ui;

    QTimer runTimer;


private:
    uint8_t decksNum;
    uint8_t sKengNum;
    uint8_t hKengNum;
    int8_t otherNum;

    double  sKengExpect[DRAW_CASE];
    double  hKengExpect[DRAW_CASE];
    double  SingExpect[DRAW_CASE];
    double  DoubExpect[DRAW_CASE];
    double  TribExpect[DRAW_CASE];
    double  KegRateExpan0[DRAW_CASE];
    double  KegRateExpan1[DRAW_CASE];
    double  KegRateExpan2[DRAW_CASE];
    double  uselessNum[DRAW_CASE];

    double  SingleP[CORE_GROUP_MAX][DEFAULT_HANDS+DRAW_CASE];
    double  DoubleP[CORE_GROUP_MAX][DEFAULT_HANDS+DRAW_CASE];
    double  TribleP[CORE_GROUP_MAX][DEFAULT_HANDS+DRAW_CASE];

    coreCard_t coreCard[CORE_GROUP_MAX];
};
#endif // MAINWINDOW_H
