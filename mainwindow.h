#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QAxWidget>
#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_actionOpen_triggered();

    void on_actionProcess_triggered();

    void on_pbPreview_clicked();

    void on_tableWidget_cellChanged(int row, int column);

    void on_pbSave_clicked();

    void on_actionClose_triggered();

private:
    Ui::MainWindow *ui;

    QString m_templateFile;
    QAxWidget* m_word01 = nullptr;
    QAxObject *m_normalStyle = nullptr;

    QHash<int, QPair<QAxObject *, QString> > m_replaces;

    void queryReplacementsInDoc();

};
#endif // MAINWINDOW_H
