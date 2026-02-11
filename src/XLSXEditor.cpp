#include "XLSXEditor.hpp"
#include "DataItem.hpp"
#include "ui_XLSXEditor.h"
#include <QGridLayout>
#include <QMessageBox>
#include <QFile>
#include <QDir>

XLSXEditor::XLSXEditor(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::XLSXEditor),
    m_wrapper(nullptr),
    m_sheetIndex(-1)
{
    ui->setupUi(this);
    connect(ui->btnSave, &QPushButton::clicked, this, &XLSXEditor::on_btnSave_clicked);
    connect(ui->btnRestore, &QPushButton::clicked, this, &XLSXEditor::on_btnRestore_clicked);
}

XLSXEditor::~XLSXEditor()
{
    delete ui;
    if (m_wrapper) {
        m_wrapper->close();
        delete m_wrapper;
    }
    for (auto item : m_dataItems) {
        delete item;
    }
}

void XLSXEditor::loadXLSX(const QString &filePath, const QString &sheetName, const QString &range)
{
    m_filePath = filePath;
    m_sheetName = sheetName;
    m_range = range;

    if (m_wrapper) {
        m_wrapper->close();
        delete m_wrapper;
    }
    m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
    if (!m_wrapper->open(filePath.toStdString())) {
        // handle error
        return;
    }

    // Find sheet index
    m_sheetIndex = -1;
    for (unsigned int i = 0; i < m_wrapper->sheetCount(); ++i) {
        if (m_wrapper->sheetName(i) == sheetName.toStdString()) {
            m_sheetIndex = static_cast<int>(i);
            break;
        }
    }
    if (m_sheetIndex < 0) {
        // handle error
        return;
    }

    // Calculate total cells
    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);
    int totalCells = (endRow - startRow + 1) * (endCol - startCol + 1);

    QProgressDialog progress(tr("Loading data..."), tr("Cancel"), 0, totalCells, this);
    progress.setWindowModality(Qt::WindowModal);

    loadData(progress);
    displayData();
}

void XLSXEditor::parseRange(const QString &range, int &startRow, int &startCol, int &endRow, int &endCol)
{
    QStringList parts = range.split(',');
    if (parts.size() == 2) {
        QString colPart = parts[0];
        QString rowPart = parts[1];
        QStringList colRange = colPart.split(':');
        if (colRange.size() == 2) {
            startCol = colToNum(colRange[0]);
            endCol = colToNum(colRange[1]);
        }
        QStringList rowRange = rowPart.split(':');
        if (rowRange.size() == 2) {
            startRow = rowRange[0].toInt();
            endRow = rowRange[1].toInt();
        }
    }
}

int XLSXEditor::colToNum(const QString &col)
{
    int num = 0;
    for (QChar c : col) {
        num = num * 26 + (c.toUpper().unicode() - 'A' + 1);
    }
    return num;
}

QString XLSXEditor::numToCol(int num)
{
    QString col;
    while (num > 0) {
        num--;
        col.prepend(QChar('A' + (num % 26)));
        num /= 26;
    }
    return col;
}

void XLSXEditor::loadData(QProgressDialog &progress)
{
    m_data.clear();
    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);

    int current = 0;
    // For each cell in range, check if it has picture
    for (int row = startRow; row <= endRow; ++row) {
        for (int col = startCol; col <= endCol; ++col) {
            progress.setValue(++current);
            if (progress.wasCanceled()) {
                return;
            }
            QString cellRef = numToCol(col) + QString::number(row);
            auto rawOpt = m_wrapper->getPictureRaw(static_cast<unsigned int>(m_sheetIndex), cellRef.toStdString());
            if (rawOpt.has_value()) {
                const std::vector<uint8_t> &bytes = rawOpt.value();
                QImage image = QImage::fromData(bytes.data(), bytes.size());
                // Get description from cell below
                QString descCell = numToCol(col) + QString::number(row + 1);
                auto descOpt = m_wrapper->getCellValue(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString());
                QString desc = descOpt.has_value() ? QString::fromStdString(descOpt.value()) : "";
                m_data.append({row, col, image, desc, false});
            }
        }
    }
}

void XLSXEditor::displayData()
{
    // Clear existing
    for (auto item : m_dataItems) {
        ui->gridData->removeWidget(item);
        delete item;
    }
    m_dataItems.clear();

    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);

    QGridLayout *layout = ui->gridData;

    for (int i = 0; i < m_data.size(); ++i) {
        const auto &entry = m_data[i];
        DataItem *item = new DataItem(this);
        item->setImage(entry.image);
        item->setDescription(entry.desc);
        item->setDeleted(entry.deleted);
        item->setRowCol(entry.row, entry.col);
        int gridRow = entry.row - startRow;
        int gridCol = entry.col - startCol;
        layout->addWidget(item, gridRow, gridCol);
        connect(item, &DataItem::deleteToggled, [this, i](bool deleted) {
            m_data[i].deleted = deleted;
        });
        m_dataItems.append(item);
    }
}

void XLSXEditor::on_btnSave_clicked()
{
    saveData();
    QMessageBox::information(this, tr("Save"), tr("Data saved to XLSX."));
}

void XLSXEditor::on_btnRestore_clicked()
{
    restoreData();
    QMessageBox::information(this, tr("Restore"), tr("Data restored."));
}

void XLSXEditor::saveData()
{
    // For each data item, if deleted, set background to red
    for (int i = 0; i < m_data.size(); ++i) {
        const auto &entry = m_data[i];
        QString descCell = numToCol(entry.col) + QString::number(entry.row + 1);
        cc::neolux::utils::MiniXLSX::CellStyle cs;
        if (entry.deleted) {
            cs.backgroundColor = "#FF0000"; // red
        } else {
            cs.backgroundColor = ""; // default
        }
        m_wrapper->setCellStyle(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(), cs);
    }
    m_wrapper->save();
}

void XLSXEditor::restoreData()
{
    // Restore from cache
    displayData();
}


