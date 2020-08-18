#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFileDialog>
#include <QDebug>
#include <QAxObject>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    if (m_word01) {
        QAxObject *wApp = m_word01->querySubObject("Application");
        wApp->dynamicCall("Quit (void)");
    }

    delete ui;
}


void MainWindow::on_actionOpen_triggered()
{
    m_templateFile = QFileDialog::getOpenFileName(this, tr("Open Document Template"), "D:/", tr("Document Files (*.docx *.doc)"));
    qDebug() << "open template:" << m_templateFile;

    m_word01 = new QAxWidget("Word.Application", this);
//    m_word01->querySubObject("Documents")->querySubObject("Open(const QString&, bool)", m_templateFile, true);
    m_word01->show();
    m_word01->setControl(m_templateFile);

    ui->scrollAreaWidgetContents_1->layout()->addWidget(m_word01);

    ui->tableWidget->setRowCount(0);
//    queryReplacementsInDoc();
}

void MainWindow::queryReplacementsInDoc()
{
    if (m_templateFile.isEmpty() || !m_word01)
        return;

    QAxObject *wApp = m_word01->querySubObject("Application");
    wApp->querySubObject("Documents")->querySubObject("Item(int)", 1)->dynamicCall("Activate()");
    QAxObject *doc = wApp->querySubObject("ActiveDocument");

    // check style
//    QAxObject *styles = doc->querySubObject("Styles");
//    int styleCount = styles->dynamicCall("Count()").toInt();
//    for (int i = 1; i <= styleCount; i++){
//        QAxObject *st = styles->querySubObject("Item(int)", i);
//        QString name = st->dynamicCall("NameLocal()").toString();
//        qDebug() << "style" << name;
//    }

    // query all words
    QAxObject *words = doc->querySubObject("Words");
    int countWord = words->dynamicCall("Count()").toInt();
    for (int i = 1; i <= countWord; i++){
        QAxObject *range = words->querySubObject("Item(int)", i);
        QAxObject *style = range->querySubObject("Style()");
        QString w = range->dynamicCall("Text()").toString();
        QString st = style->dynamicCall("NameLocal()").toString();
//        qDebug() << w << st;

        if (st.toLower().startsWith("subtitle")) {
            qDebug() << "FOUND word to replace:" << w;
            int r = ui->tableWidget->rowCount();
            ui->tableWidget->insertRow(r);
            ui->tableWidget->setItem(r, 0, new QTableWidgetItem(w));
            ui->tableWidget->setItem(r, 1, new QTableWidgetItem(w.toUpper()));

            m_replaces.insert(r, QPair<QAxObject *, QString>(range, w.toUpper()));      // test
        }
        if (!m_normalStyle && st.toLower() == "normal") {
            m_normalStyle = style;
            qDebug() << "FOUND normal style";
        }
    }
}

void MainWindow::on_actionProcess_triggered()
{
    queryReplacementsInDoc();
}

void MainWindow::on_pbPreview_clicked()
{
    if (!m_word01)
        return;

    QAxObject *wApp = m_word01->querySubObject("Application");

    for (auto itr = m_replaces.begin(); itr != m_replaces.end(); itr++) {
//        qDebug() << "before replaced:" << itr.value().first->dynamicCall("Text()").toString() << "->" << itr.value().second;
        itr.value().first->setProperty("Text", itr.value().second);
        itr.value().first->dynamicCall("Select()");
        QAxObject *selection = wApp->querySubObject("Selection");
        selection->dynamicCall("ClearFormatting ()");

//        itr.value().first->setProperty("Style", m_normalStyle->dynamicCall("NameLocal()"));
//        qDebug() << "after replaced:" << itr.value().first->dynamicCall("Text()").toString();

        QAxObject *style = itr.value().first->querySubObject("Style()");
        qDebug() << "change style:" << style->dynamicCall("NameLocal()").toString();
    }
}

void MainWindow::on_tableWidget_cellChanged(int row, int column)
{
    if (row >= 0 && row < ui->tableWidget->rowCount() && column == 1) {
        auto itr = m_replaces.find(row);
        if (itr != m_replaces.end()) {
            itr.value().second = ui->tableWidget->item(row, column)->text().trimmed() + " ";
            qDebug() << "update value:" << row << column << itr.value().second;
        }
    }
}

void MainWindow::on_pbSave_clicked()
{
    QString path = ui->lineEdit->text().trimmed();
    // check valid path?

    QAxObject *wApp = m_word01->querySubObject("Application");
    QAxObject *doc = wApp->querySubObject("ActiveDocument");
    doc->dynamicCall("SaveAs (const QString&)", path);
}

void MainWindow::on_actionClose_triggered()
{
    if (m_word01) {
        QAxObject *wApp = m_word01->querySubObject("Application");
        wApp->dynamicCall("Quit ()");
    }
}
