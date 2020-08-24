#ifndef DOCXREPORT_H
#define DOCXREPORT_H

#include <QAxWidget>
#include <QComboBox>
#include <QMainWindow>

namespace Ui {
class DocxReport;
}

class DocxReport : public QMainWindow
{
    Q_OBJECT

public:
    explicit DocxReport(QWidget *parent = nullptr);
    ~DocxReport();

private slots:
    void on_actionOpen_triggered();

    void on_pbProcess_clicked();

    void on_pbPreview_clicked();

    void on_actionClose_triggered();

    void on_pbSave_clicked();

private:
    Ui::DocxReport *ui;

    QStringList m_ranks;

    QString m_templateFile;
    QAxWidget *m_wWidget = nullptr;
    QAxObject *m_wApp = nullptr;
    QAxObject *m_normalStyle = nullptr;

    QStringList m_learners;
    QHash<int, QComboBox *> m_cbLearnerRanks;
    QHash<QString, QAxObject *> m_replaces;


    void initCbRank(QComboBox *cb);
    void initComboboxes();
    QString getValueByKey(const QString &key);


    // random data
    QStringList m_lastNames;
    QStringList m_middleNames;
    QStringList m_firstNames;

    void makeSampleData();
    QString randomName();

};

#endif // DOCXREPORT_H
