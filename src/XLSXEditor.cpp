#include "cc/neolux/fem/xlsxeditor/XLSXEditor.hpp"
#include "cc/neolux/fem/xlsxeditor/DataItem.hpp"
#include "ui_XLSXEditor.h"
#include <QGridLayout>
#include <QMessageBox>
#include <QFile>
#include <QDialog>
#include <QLabel>
#include <QVBoxLayout>
#include <QProgressBar>
#include <QCoreApplication>
#include <QTimer>
#include <QDebug>
#include <fstream>
#include <utility>

namespace {
bool splitCellRef(const QString &ref, QString &colPart, QString &rowPart)
{
    colPart.clear();
    rowPart.clear();
    for (QChar c : ref.trimmed()) {
        if (c.isLetter()) {
            colPart.append(c.toUpper());
        } else if (c.isDigit()) {
            rowPart.append(c);
        }
    }
    return !colPart.isEmpty() && !rowPart.isEmpty();
}
}

using namespace cc::neolux::fem::xlsxeditor;

XLSXEditor::XLSXEditor(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::XLSXEditor),
    m_wrapper(nullptr),
    m_sheetIndex(-1)
{
    ui->setupUi(this);
    ui->progressBar->setVisible(false);
}

QString XLSXEditor::cellKey(int row, int col) const
{
    return QString::number(row) + QLatin1Char(':') + QString::number(col);
}

XLSXEditor::~XLSXEditor()
{
    clearDataItems();
    delete ui;
    if (m_wrapper) {
        m_wrapper->close();
        delete m_wrapper;
    }
    if (m_pictureReader.isOpen()) {
        m_pictureReader.close();
    }
}

void XLSXEditor::loadXLSX(const QString &filePath, const QString &sheetName, const QString &range)
{
    resetState();
    m_filePath = filePath;
    m_sheetName = sheetName;
    m_range = range;

    m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
    if (!m_wrapper->open(filePath.toStdString())) {
        QMessageBox::critical(this, tr("Error"), tr("Failed to open XLSX file."));
        ui->progressBar->setVisible(false);
        return;
    }

    if (!m_pictureReader.open(filePath.toStdString())) {
        QMessageBox::critical(this, tr("Error"), tr("Failed to prepare picture reader."));
        m_wrapper->close();
        ui->progressBar->setVisible(false);
        return;
    }

    // 查找表索引
    m_sheetIndex = -1;
    for (unsigned int i = 0; i < m_wrapper->sheetCount(); ++i) {
        if (m_wrapper->sheetName(i) == sheetName.toStdString()) {
            m_sheetIndex = static_cast<int>(i);
            break;
        }
    }
    if (m_sheetIndex < 0) {
        QMessageBox::critical(this, tr("Error"), tr("Sheet not found: %1").arg(sheetName));
        ui->progressBar->setVisible(false);
        return;
    }

    // 计算总单元格数
    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);
    int totalCells = (endRow - startRow + 1) * (endCol - startCol + 1);

    ui->progressBar->setMinimum(0);
    ui->progressBar->setMaximum(totalCells);
    ui->progressBar->setValue(0);
    ui->progressBar->setVisible(true);
    QCoreApplication::processEvents();

    loadData(*ui->progressBar);
    displayData();
    ui->progressBar->setVisible(false);
}

void XLSXEditor::parseRange(const QString &range, int &startRow, int &startCol, int &endRow, int &endCol)
{
    startRow = 0;
    startCol = 0;
    endRow = 0;
    endCol = 0;

    QString trimmed = range.trimmed();
    if (trimmed.contains(',')) {
        // 格式：A:C,1:10
        QStringList parts = trimmed.split(',');
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
    } else if (trimmed.contains(':')) {
        // 格式：A1:C10
        QStringList parts = trimmed.split(':');
        if (parts.size() == 2) {
            QString startColPart;
            QString startRowPart;
            QString endColPart;
            QString endRowPart;
            if (splitCellRef(parts[0], startColPart, startRowPart) &&
                splitCellRef(parts[1], endColPart, endRowPart)) {
                startCol = colToNum(startColPart);
                endCol = colToNum(endColPart);
                startRow = startRowPart.toInt();
                endRow = endRowPart.toInt();
            }
        }
    }

    if (startRow > endRow) {
        std::swap(startRow, endRow);
    }
    if (startCol > endCol) {
        std::swap(startCol, endCol);
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

void XLSXEditor::loadData(QProgressBar &progressBar)
{
    m_data.clear();
    m_indexByCell.clear();
    m_itemByCell.clear();
    m_dirtyCells.clear();
    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);

    // 通过图片读取器获取表内图片
    auto allPictures = m_pictureReader.getSheetPictures(static_cast<unsigned int>(m_sheetIndex));
    std::string tempDir = m_pictureReader.getTempDir();
    if (tempDir.empty()) {
        qWarning() << "Failed to extract XLSX temporary files.";
        progressBar.setVisible(false);
        return;
    }

    int totalPictures = 0;
    for (const auto &pic : allPictures) {
        if (pic.rowNum >= startRow && pic.rowNum <= endRow &&
            pic.colNum >= startCol && pic.colNum <= endCol) {
            ++totalPictures;
        }
    }
    qInfo() << "loadData totalPictures" << totalPictures
            << "range" << startRow << startCol << endRow << endCol;
    progressBar.setMaximum(totalPictures);

    int current = 0;
    for (const auto &pic : allPictures) {
        if (pic.rowNum < startRow || pic.rowNum > endRow ||
            pic.colNum < startCol || pic.colNum > endCol) {
            continue;
        }

        progressBar.setValue(++current);
        QCoreApplication::processEvents();

        std::string imagePath = tempDir + "/xl/" + pic.relativePath;
        std::ifstream file(imagePath, std::ios::binary);
        QImage image;
        if (file) {
            std::vector<uint8_t> bytes((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
            image = QImage::fromData(bytes.data(), bytes.size());
            if (image.isNull()) {
                // QMessageBox::warning(this, tr("Warning"), tr("Failed to load image from %1").arg(QString::fromStdString(imagePath)));
                qWarning() << "Failed to load image from" << QString::fromStdString(imagePath);
            }
        } else {
            qWarning() << "Image file not found:" << QString::fromStdString(imagePath);
        }

        QString descCell = numToCol(pic.colNum) + QString::number(pic.rowNum + 1);
        auto descOpt = m_wrapper->getCellValue(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString());
        QString desc = descOpt.has_value() ? QString::fromStdString(descOpt.value()) : "";
        m_data.append({pic.rowNum, pic.colNum, image, desc, false});
        m_indexByCell.insert(cellKey(pic.rowNum, pic.colNum), m_data.size() - 1);
    }
}

void XLSXEditor::clearDataItems()
{
    if (!ui) {
        return;
    }
    for (auto item : m_dataItems) {
        ui->gridData->removeWidget(item);
        delete item;
    }
    m_dataItems.clear();
    m_itemByCell.clear();
}

void XLSXEditor::resetState()
{
    clearDataItems();
    m_data.clear();
    m_indexByCell.clear();
    m_dirtyCells.clear();
    m_sheetIndex = -1;
    if (m_wrapper) {
        m_wrapper->close();
        delete m_wrapper;
        m_wrapper = nullptr;
    }
    if (m_pictureReader.isOpen()) {
        m_pictureReader.close();
    }
    if (ui && ui->progressBar) {
        ui->progressBar->setValue(0);
        ui->progressBar->setVisible(false);
    }
}

void XLSXEditor::displayData()
{
    // 清理旧组件
    clearDataItems();

    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);

    QGridLayout *layout = ui->gridData;
    layout->setSpacing(0);
    layout->setContentsMargins(0, 0, 0, 0);

    for (int i = 0; i < m_data.size(); ++i) {
        const auto &entry = m_data[i];
        if (!entry.image.isNull()) {
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
                m_dirtyCells.insert(cellKey(m_data[i].row, m_data[i].col));
            });
            connect(item, &DataItem::imageClicked, this, &XLSXEditor::showImageDialog);
            m_dataItems.append(item);
            m_itemByCell.insert(cellKey(entry.row, entry.col), item);
        }
    }
}

void XLSXEditor::on_btnSave_clicked()
{
    if (saveData()) {
        QMessageBox::information(this, tr("Save"), tr("Data saved to XLSX."));
    } else {
        QMessageBox::critical(this, tr("Save Error"), tr("Failed to save data to XLSX."));
    }
}

void XLSXEditor::on_btnRestore_clicked()
{
    restoreData();
    QMessageBox::information(this, tr("Restore"), tr("Data restored."));
}

bool XLSXEditor::saveData()
{
    // 方案一：全量覆盖保存（当前使用）
    // 保存时覆盖所有数据点的标记状态
    for (const auto &entry : m_data) {
        QString picCell = numToCol(entry.col) + QString::number(entry.row);
        QString descCell = numToCol(entry.col) + QString::number(entry.row + 1);
        cc::neolux::utils::MiniXLSX::CellStyle cs;
        cs.backgroundColor = entry.deleted ? "#FF0000" : "";
        if (entry.deleted) {
            m_wrapper->setCellValue(static_cast<unsigned int>(m_sheetIndex), picCell.toStdString(), "[D]");
        } else {
            m_wrapper->setCellValue(static_cast<unsigned int>(m_sheetIndex), picCell.toStdString(), "");
        }
        m_wrapper->setCellStyle(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(), cs);
    }

    // 方案二：仅保存有变更的单元格（备用）
    // for (const QString &key : std::as_const(m_dirtyCells)) {
    //     auto it = m_indexByCell.find(key);
    //     if (it == m_indexByCell.end()) {
    //         continue;
    //     }
    //     const auto &entry = m_data[it.value()];
    //     QString picCell = numToCol(entry.col) + QString::number(entry.row);
    //     QString descCell = numToCol(entry.col) + QString::number(entry.row + 1);
    //     cc::neolux::utils::MiniXLSX::CellStyle cs;
    //     cs.backgroundColor = entry.deleted ? "#FF0000" : "";
    //     if (entry.deleted) {
    //         m_wrapper->setCellValue(static_cast<unsigned int>(m_sheetIndex), picCell.toStdString(), "[D]");
    //     } else {
    //         m_wrapper->setCellValue(static_cast<unsigned int>(m_sheetIndex), picCell.toStdString(), "");
    //     }
    //     m_wrapper->setCellStyle(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(), cs);
    // }

    bool saved = m_wrapper->save();
    if (saved) {
        m_dirtyCells.clear();
    }
    return saved;
}

void XLSXEditor::showImageDialog(int row, int col)
{
    for (const auto &entry : m_data) {
        if (entry.row == row && entry.col == col && !entry.image.isNull()) {
            QDialog dialog(this);
            dialog.setWindowTitle(tr("Image Preview"));
            QLabel *label = new QLabel(&dialog);
            label->setPixmap(QPixmap::fromImage(entry.image));
            QVBoxLayout *layout = new QVBoxLayout(&dialog);
            layout->addWidget(label);
            dialog.setLayout(layout);
            dialog.exec();
            break;
        }
    }
}

void XLSXEditor::restoreData()
{
    for (const QString &key : std::as_const(m_dirtyCells)) {
        auto indexIt = m_indexByCell.find(key);
        if (indexIt == m_indexByCell.end()) {
            continue;
        }
        int index = indexIt.value();
        m_data[index].deleted = false;

        auto itemIt = m_itemByCell.find(key);
        if (itemIt != m_itemByCell.end() && itemIt.value()) {
            itemIt.value()->setDescription(m_data[index].desc);
            itemIt.value()->setDeleted(false);
        }

        QString picCell = numToCol(m_data[index].col) + QString::number(m_data[index].row);
        QString descCell = numToCol(m_data[index].col) + QString::number(m_data[index].row + 1);
        cc::neolux::utils::MiniXLSX::CellStyle cs;
        cs.backgroundColor = "";
        m_wrapper->setCellValue(static_cast<unsigned int>(m_sheetIndex), picCell.toStdString(), "");
        m_wrapper->setCellStyle(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(), cs);
    }
    m_wrapper->save();
    m_dirtyCells.clear();
}


