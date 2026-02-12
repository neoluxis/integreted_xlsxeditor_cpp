#include "cc/neolux/fem/xlsxeditor/DataItem.hpp"
#include "ui_DataItem.h"
#include <QPixmap>
#include <QLabel>
#include <QDebug>

using namespace cc::neolux::fem::xlsxeditor;

DataItem::DataItem(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::DataItem),
    m_deleted(false),
    m_row(-1),
    m_col(-1)
{
    ui->setupUi(this);
    setStyleSheet("border: 2px solid gray;");
    connect(ui->btnImage, &QPushButton::pressed, this, [this]() { emit imageClicked(m_row, m_col); });
    ui->lnData->setReadOnly(true); // 数据只读，防止误修改
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
    // qInfo() << "DataItem setDeleted" << deleted << "row" << m_row << "col" << m_col;
    if (deleted) {
        ui->btnMarkDel->setText(tr("Restore"));
        // 待办：XLSX 中写入红色背景；UI 仅调整样式
        ui->lnData->setStyleSheet("background-color: red;");
    } else {
        ui->btnMarkDel->setText(tr("Delete"));
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
    qInfo() << "DataItem delete clicked" << "row" << m_row << "col" << m_col;
    setDeleted(!m_deleted);
    emit deleteToggled(m_deleted);
}