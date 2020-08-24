#include "docxreport.h"
#include "ui_docxreport.h"

#include <QAxObject>
#include <QDebug>
#include <QFileDialog>
#include <QMessageBox>
#include <QRandomGenerator>

DocxReport::DocxReport(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::DocxReport)
{
    ui->setupUi(this);

    m_ranks << tr("Thiếu tá") << tr("Trung tá") << tr("Thượng tá") << tr("Đại tá")
            << tr("Thiếu tướng") << tr("Trung tướng") << tr("Thượng tướng") << tr("Đại tướng");

    m_learners << "pilot01" << "pilot02" << "lead" << "assistant"
               << "schNav" << "k4Nav" << "k6Nav" << "k7Nav" << "k10Nav";

    initComboboxes();


    // sample
    m_firstNames << "A" << "B" << "C" << "D" << "X" << "Y" << "Z";
    m_middleNames << "Văn" << "Đình" << "Nguyên" << "Đức";
    m_lastNames << "Nguyễn" << "Trần" << "Phạm" << "Lê";
    makeSampleData();
}

DocxReport::~DocxReport()
{
    delete ui;
}

void DocxReport::on_actionOpen_triggered()
{
    m_templateFile = QFileDialog::getOpenFileName(this, tr("Open Template"), "D:/", tr("Document Files (*.docx *.doc)"));
    if (m_templateFile.isEmpty())
        return;

    qDebug() << "open template:" << m_templateFile;

    m_wWidget = new QAxWidget("Word.Application", this);
    m_wWidget->show();
    m_wWidget->setControl(m_templateFile);

    ui->scrollAreaWidgetContents->layout()->addWidget(m_wWidget);
    m_wApp = m_wWidget->querySubObject("Application");
    m_wApp->querySubObject("Documents")->querySubObject("Item(int)", 1)->dynamicCall("Activate()");
    qDebug() << "Done!";
}

void DocxReport::initCbRank(QComboBox *cb)
{
    if (cb) {
        for (int i = 0; i < m_ranks.size(); ++i) {
            cb->addItem(m_ranks.at(i), i);
        }
    }
}

void DocxReport::initComboboxes()
{
    initCbRank(ui->comboBox);

    for (int i = 0; i < ui->tableWidget->rowCount(); ++i) {
        QComboBox *cb = new QComboBox;
        initCbRank(cb);
        cb->setCurrentIndex(QRandomGenerator::global()->generate() % cb->count());
        ui->tableWidget->setCellWidget(i, 1, cb);
        m_cbLearnerRanks.insert(i, cb);
    }
}

QString DocxReport::getValueByKey(const QString &key)
{
    if (key == "place") {
        return ui->lePlace->text().trimmed();
    }
    if (key == "exercise") {
        return ui->leExercise->text().trimmed();
    }
    if (key == "starttime") {
        return ui->dteStart->dateTime().toString("dd/MM/yyyy hh:mm");
    }
    if (key == "endtime") {
        return ui->dteEnd->dateTime().toString("dd/MM/yyyy hh:mm");
    }
    if (key == "note") {
        return ui->plainTextEdit->document()->toPlainText().trimmed();
    }
    if (key.startsWith("teacher")) {
        if (key.endsWith("name")) {
            return ui->leTeacherName->text().trimmed();
        }
        if (key.endsWith("rank")) {
            return ui->comboBox->currentText();
        }
        if (key.endsWith("func")) {
            return ui->leTeacherFunc->text().trimmed();
        }
    }

    for (int i = 0; i < m_learners.size(); ++i) {
        if (key.startsWith(m_learners.at(i).toLower())) {
            if (key.endsWith("name")) {
                return ui->tableWidget->item(i, 0)->text();
            }
            if (key.endsWith("rank")) {
                return m_cbLearnerRanks.value(i)->currentText();
            }
            if (key.endsWith("func")) {
                return ui->tableWidget->item(i, 2)->text();
            }
            if (key.endsWith("result")) {
                return ui->tableWidget->item(i, 3)->text();
            }
        }
    }
    return " ";
}

void DocxReport::makeSampleData()
{
    for (int i = 0; i < ui->tableWidget->rowCount(); ++i) {
        ui->tableWidget->setItem(i, 0, new QTableWidgetItem(randomName()));
        ui->tableWidget->setItem(i, 2, new QTableWidgetItem(tr("Không")));
        int point = QRandomGenerator::global()->generate() % 5;
        ui->tableWidget->setItem(i, 3, new QTableWidgetItem(QString("%1/10").arg(5 + point)));
    }
}

QString DocxReport::randomName()
{
    int i = QRandomGenerator::global()->generate() % m_lastNames.size();
    int j = QRandomGenerator::global()->generate() % m_middleNames.size();
    int k = QRandomGenerator::global()->generate() % m_firstNames.size();

    return QString("%1 %2 %3").arg(m_lastNames.at(i)).arg(m_middleNames.at(j)).arg(m_firstNames.at(k));
}

void DocxReport::on_pbProcess_clicked()
{
    if (m_templateFile.isEmpty() || !m_wWidget || !m_wApp) {
        qCritical() << "INVALID file";
        return;
    }

//    QAxObject *wApp = m_wWidget->querySubObject("Application");
//    wApp->querySubObject("Documents")->querySubObject("Item(int)", 1)->dynamicCall("Activate()");
    QAxObject *doc = m_wApp->querySubObject("ActiveDocument");

    QAxObject *words = doc->querySubObject("Words");
    int countWord = words->dynamicCall("Count()").toInt();
    for (int i = 1; i <= countWord; i++){
        QAxObject *range = words->querySubObject("Item(int)", i);
        QAxObject *style = range->querySubObject("Style()");
        QString w = range->dynamicCall("Text()").toString();
        QString st = style->dynamicCall("NameLocal()").toString();

        if (st.toLower().startsWith("subtitle") && !w.trimmed().isEmpty()) {
            qDebug() << "FOUND word to replace:" << w;
            m_replaces.insert(w.trimmed().toLower(), range);
        }
        if (!m_normalStyle && st.toLower() == "normal") {
            m_normalStyle = style;
            qDebug() << "FOUND normal style";
        }
    }
    qDebug() << "DONE!!!";
    QMessageBox::information(this, "Thông báo", "Đã check template!");
}

void DocxReport::on_pbPreview_clicked()
{
    if (!m_wWidget)
        return;

    QAxObject *wApp = m_wWidget->querySubObject("Application");
    for (auto itr = m_replaces.begin(); itr != m_replaces.end(); itr++) {
        QString v = getValueByKey(itr.key());
        qDebug() << "get value by key" << itr.key() << ":" << v;

        itr.value()->setProperty("Text", v);
        itr.value()->dynamicCall("Select()");
        QAxObject *selection = wApp->querySubObject("Selection");
        selection->dynamicCall("ClearFormatting ()");

        QAxObject *style = itr.value()->querySubObject("Style()");
    }
}

void DocxReport::on_actionClose_triggered()
{
    if (m_wWidget && m_wApp) {
        QAxObject *doc = m_wApp->querySubObject("Documents");
        if (doc) {
            qDebug() << "close doc?";
            doc->dynamicCall("Close (0)");
        }
        m_wApp->dynamicCall("Quit (void)");
        m_wWidget->clear();
//        ui->scrollAreaWidgetContents->layout()->removeWidget(m_wWidget);

        delete m_wWidget;
        m_wWidget = nullptr;
    }
}

void DocxReport::on_pbSave_clicked()
{
    if (!m_wWidget || !m_wApp) {
        QMessageBox::warning(this, "Cảnh báo", "Chưa mở file word!");
    }

    QString save = QFileDialog::getSaveFileName(this, tr("Save Report"), "D:/", tr("Document Files (*.docx *.doc)"));
    if (save.isEmpty())
        return;

    QAxObject *wApp = m_wWidget->querySubObject("Application");
    QAxObject *doc = wApp->querySubObject("ActiveDocument");
    doc->dynamicCall("SaveAs (const QString&)", save);
}
