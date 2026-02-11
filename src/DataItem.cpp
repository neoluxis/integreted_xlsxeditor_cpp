#include "DataItem.hpp"
#include "ui_DataItem.h"
#include <QPixmap>
#include <QLabel>

DataItem::DataItem(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::DataItem),
    m_deleted(false),
    m_row(-1),
    m_col(-1)
{
    ui->setupUi(this);
    setStyleSheet("border: 1px solid lightgray;");
    connect(ui->btnMarkDel, &QPushButton::clicked, this, &DataItem::on_btnMarkDel_clicked);
    connect(ui->btnImage, &QPushButton::clicked, this, [this]() { emit imageClicked(m_row, m_col); });
}

DataItem::~DataItem()
{
    delete ui;
}

void DataItem::setImage(const QImage &image)
{
    if (!image.isNull()) {
        QPixmap pixmap = QPixmap::fromImage(image);
        ui->btnImage->setIcon(QIcon(pixmap.scaled(70, 70, Qt::KeepAspectRatio)));
    } else {
        ui->btnImage->setIcon(QIcon());
    }
}

void DataItem::setDescription(const QString &desc)
{
    ui->lnData->setText(desc);
}

QString DataItem::getDescription() const
{
    return ui->lnData->text();
}

void DataItem::setDeleted(bool deleted)
{
    m_deleted = deleted;
    if (deleted) {
        ui->lnData->setText("[Deleted] " + ui->lnData->text());
        // TODO: set background color to red for XLSX, but for UI, maybe change style
        ui->lnData->setStyleSheet("background-color: red;");
    } else {
        QString text = ui->lnData->text();
        if (text.startsWith("[Deleted] ")) {
            text = text.mid(10);
        }
        ui->lnData->setText(text);
        ui->lnData->setStyleSheet("");
    }
}

bool DataItem::isDeleted() const
{
    return m_deleted;
}

void DataItem::setRowCol(int row, int col)
{
    m_row = row;
    m_col = col;
}

int DataItem::getRow() const
{
    return m_row;
}

int DataItem::getCol() const
{
    return m_col;
}

void DataItem::on_btnMarkDel_clicked()
{
    setDeleted(!m_deleted);
    emit deleteToggled(m_deleted);
}