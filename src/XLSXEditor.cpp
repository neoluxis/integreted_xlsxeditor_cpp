#include "cc/neolux/fem/xlsxeditor/XLSXEditor.hpp"

#include <QCheckBox>
#include <QCoreApplication>
#include <QDebug>
#include <QDialog>
#include <QDir>
#include <QEvent>
#include <QFile>
#include <QFileInfo>
#include <QGridLayout>
#include <QLabel>
#include <QMessageBox>
#include <QProgressBar>
#include <QTimer>
#include <QVBoxLayout>
#include <QWheelEvent>
#include <algorithm>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <pugixml.hpp>
#include <set>
#include <unordered_set>
#include <utility>

#include "cc/neolux/fem/xlsxeditor/DataItem.hpp"
#include "cc/neolux/utils/KFZippa/kfzippa.hpp"
#include "ui_XLSXEditor.h"

namespace {
constexpr int kBaseItemWidth = 70;
constexpr int kBaseItemHeight = 90;
constexpr int kBaseHeaderColWidth = 90;
constexpr int kBaseHeaderRowHeight = 24;
constexpr int kGridSpacing = 0;
constexpr bool kEnableSaveProgress = true;

QString tx(const char* sourceText) {
    return QCoreApplication::translate("XLSXEditor", sourceText);
}

bool splitCellRef(const QString& ref, QString& colPart, QString& rowPart) {
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
}  // namespace

using namespace cc::neolux::fem::xlsxeditor;

XLSXEditor::XLSXEditor(QWidget* parent, bool dry_run)
    : QWidget(parent),
      ui(new Ui::XLSXEditor),
      m_wrapper(nullptr),
      m_sheetIndex(-1),
      m_enableSaveProgress(kEnableSaveProgress),
      m_dryRun(dry_run),
      m_previewOnly(false),
      m_itemScale(1.0),
      m_syncingSelectAll(false) {
    ui->setupUi(this);
    ui->progressBar->setVisible(false);
    ui->scrollArea->setWidgetResizable(false);
    ui->scrollArea->setAlignment(Qt::AlignLeft | Qt::AlignTop);
    ui->scrollArea->viewport()->installEventFilter(this);
    syncPreviewButtonText();
}

QString XLSXEditor::cellKey(int row, int col) const {
    return QString::number(row) + QLatin1Char(':') + QString::number(col);
}

XLSXEditor::~XLSXEditor() {
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

void XLSXEditor::loadXLSX(const QString& filePath, const QString& sheetName, const QString& range) {
    resetState();
    m_filePath = filePath;
    m_sheetName = sheetName;
    m_range = range;

    m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
    if (!m_wrapper->open(filePath.toStdString())) {
        QMessageBox::critical(this, tx("Error"), tx("Failed to open XLSX file."));
        ui->progressBar->setVisible(false);
        return;
    }

    if (!m_pictureReader.open(filePath.toStdString())) {
        QMessageBox::critical(this, tx("Error"), tx("Failed to prepare picture reader."));
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
        QMessageBox::critical(this, tx("Error"), tx("Sheet not found: %1").arg(sheetName));
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
    displayData(false);
    ui->progressBar->setVisible(false);
}

void XLSXEditor::setDryRun(bool dry_run) {
    m_dryRun = dry_run;
}

bool XLSXEditor::isDryRun() const {
    return m_dryRun;
}

void XLSXEditor::parseRange(const QString& range, int& startRow, int& startCol, int& endRow,
                            int& endCol) {
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

int XLSXEditor::colToNum(const QString& col) {
    int num = 0;
    for (QChar c : col) {
        num = num * 26 + (c.toUpper().unicode() - 'A' + 1);
    }
    return num;
}

QString XLSXEditor::numToCol(int num) {
    QString col;
    while (num > 0) {
        num--;
        col.prepend(QChar('A' + (num % 26)));
        num /= 26;
    }
    return col;
}

QString XLSXEditor::readCellText(int row, int col) {
    if (row <= 0 || col <= 0 || m_wrapper == nullptr || m_sheetIndex < 0) {
        return "";
    }

    const QString cell = numToCol(col) + QString::number(row);
    auto cellOpt =
        m_wrapper->getCellValue(static_cast<unsigned int>(m_sheetIndex), cell.toStdString());
    return cellOpt.has_value() ? QString::fromStdString(cellOpt.value()).trimmed() : "";
}

void XLSXEditor::loadData(QProgressBar& progressBar) {
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
    for (const auto& pic : allPictures) {
        if (pic.rowNum >= startRow && pic.rowNum <= endRow && pic.colNum >= startCol &&
            pic.colNum <= endCol) {
            ++totalPictures;
        }
    }
    progressBar.setMaximum(totalPictures);

    int current = 0;
    for (const auto& pic : allPictures) {
        if (pic.rowNum < startRow || pic.rowNum > endRow || pic.colNum < startCol ||
            pic.colNum > endCol) {
            continue;
        }

        progressBar.setValue(++current);
        QCoreApplication::processEvents();

        std::string imagePath = tempDir + "/xl/" + pic.relativePath;
        std::ifstream file(imagePath, std::ios::binary);
        QImage image;
        if (file) {
            std::vector<uint8_t> bytes((std::istreambuf_iterator<char>(file)),
                                       std::istreambuf_iterator<char>());
            image = QImage::fromData(bytes.data(), bytes.size());
            if (image.isNull()) {
                // QMessageBox::warning(this, tr("Warning"), tr("Failed to load image from
                // %1").arg(QString::fromStdString(imagePath)));
                qWarning() << "Failed to load image from" << QString::fromStdString(imagePath);
            }
        } else {
            qWarning() << "Image file not found:" << QString::fromStdString(imagePath);
        }

        const QString value = readCellText(pic.rowNum + 1, pic.colNum);
        m_data.append({pic.rowNum, pic.colNum, image, value, false});
        m_indexByCell.insert(cellKey(pic.rowNum, pic.colNum), m_data.size() - 1);
    }
}

void XLSXEditor::clearDataItems() {
    if (!ui) {
        return;
    }
    for (auto item : m_dataItems) {
        ui->gridData->removeWidget(item);
        delete item;
    }
    for (auto widget : m_headerWidgets) {
        ui->gridData->removeWidget(widget);
        delete widget;
    }
    m_dataItems.clear();
    m_headerWidgets.clear();
    m_itemByCell.clear();
}

void XLSXEditor::resetState() {
    clearDataItems();
    m_data.clear();
    m_indexByCell.clear();
    m_dirtyCells.clear();
    m_previewOnly = false;
    m_itemScale = 1.0;
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
    syncPreviewButtonText();
    syncSelectAllState();
}

void XLSXEditor::displayData(bool previewOnly) {
    m_previewOnly = previewOnly;
    syncPreviewButtonText();
    // 清理旧组件
    clearDataItems();

    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);

    QGridLayout* layout = ui->gridData;
    layout->setSpacing(kGridSpacing);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setAlignment(Qt::AlignTop | Qt::AlignLeft);

    QSet<int> rowSet;
    QSet<int> colSet;
    for (const auto& entry : std::as_const(m_data)) {
        if (!entry.image.isNull()) {
            rowSet.insert(entry.row);
            colSet.insert(entry.col);
        }
    }
    QVector<int> displayRows = rowSet.values();
    QVector<int> displayCols = colSet.values();
    std::sort(displayRows.begin(), displayRows.end());
    std::sort(displayCols.begin(), displayCols.end());

    QHash<int, int> rowToGridRow;
    QHash<int, int> colToGridCol;
    for (int i = 0; i < displayRows.size(); ++i) {
        rowToGridRow.insert(displayRows[i], i + 1);
    }
    for (int i = 0; i < displayCols.size(); ++i) {
        colToGridCol.insert(displayCols[i], i + 1);
    }

    const int headerWidth = static_cast<int>(std::round(kBaseHeaderColWidth * m_itemScale));
    const int headerHeight = static_cast<int>(std::round(kBaseHeaderRowHeight * m_itemScale));

    QLabel* corner = new QLabel(this);
    corner->setText(" ");
    corner->setAlignment(Qt::AlignCenter);
    corner->setFixedSize(headerWidth, headerHeight);
    layout->addWidget(corner, 0, 0);
    m_headerWidgets.append(corner);

    for (int i = 0; i < displayCols.size(); ++i) {
        const int col = displayCols[i];
        QString header = readCellText(2, col);
        if (header.isEmpty()) {
            header = numToCol(col);
        }

        QLabel* colHeader = new QLabel(this);
        colHeader->setText(header);
        colHeader->setAlignment(Qt::AlignCenter);
        colHeader->setWordWrap(true);
        colHeader->setToolTip(tx("Double-click to toggle this column"));
        colHeader->setProperty("batchAxis", "col");
        colHeader->setProperty("batchIndex", col);
        colHeader->installEventFilter(this);
        colHeader->setFixedSize(static_cast<int>(std::round(kBaseItemWidth * m_itemScale)),
                                headerHeight);
        layout->addWidget(colHeader, 0, i + 1);
        m_headerWidgets.append(colHeader);
    }

    for (int i = 0; i < displayRows.size(); ++i) {
        const int row = displayRows[i];
        QString header = readCellText(row, 1);

        QLabel* rowHeader = new QLabel(this);
        rowHeader->setText(header);
        rowHeader->setAlignment(Qt::AlignCenter);
        rowHeader->setWordWrap(true);
        rowHeader->setToolTip(tx("Double-click to toggle this row"));
        rowHeader->setProperty("batchAxis", "row");
        rowHeader->setProperty("batchIndex", row);
        rowHeader->installEventFilter(this);
        rowHeader->setFixedSize(headerWidth,
                                static_cast<int>(std::round(kBaseItemHeight * m_itemScale)));
        layout->addWidget(rowHeader, i + 1, 0);
        m_headerWidgets.append(rowHeader);
    }

    for (int i = 0; i < m_data.size(); ++i) {
        const auto& entry = m_data[i];
        if (entry.image.isNull()) {
            continue;
        }

        DataItem* item = new DataItem(this);
        item->applyScale(m_itemScale);
        item->setImage(entry.image);
        item->setDescription(entry.desc);
        item->setDeleted(entry.deleted);
        item->setRowCol(entry.row, entry.col);
        if (!rowToGridRow.contains(entry.row)) {
            continue;
        }
        if (!colToGridCol.contains(entry.col)) {
            continue;
        }
        int gridRow = rowToGridRow.value(entry.row);
        int gridCol = colToGridCol.value(entry.col);
        layout->addWidget(item, gridRow, gridCol);
        connect(item, &DataItem::deleteToggled, [this, i](bool deleted) {
            m_data[i].deleted = deleted;
            m_dirtyCells.insert(cellKey(m_data[i].row, m_data[i].col));
            syncPreviewVisibility();
            syncSelectAllState();
        });
        connect(item, &DataItem::imageClicked, this, &XLSXEditor::showImageDialog);
        m_dataItems.append(item);
        m_itemByCell.insert(cellKey(entry.row, entry.col), item);
    }

    syncPreviewVisibility();
    syncSelectAllState();
    updateScrollWidgetSize();
}

void XLSXEditor::on_btnSave_clicked() {
    if (saveData()) {
        QMessageBox::information(this, tx("Save"),
                                 tx("Data saved to XLSX: %1").arg(m_saveFilePath));
    } else {
        QMessageBox::critical(this, tx("Save Error"), tx("Failed to save data to XLSX."));
    }
}

void XLSXEditor::on_btnPreview_clicked() {
    m_previewOnly = !m_previewOnly;
    syncPreviewButtonText();
    syncPreviewVisibility();
}

void XLSXEditor::on_chkSelectAll_stateChanged(int state) {
    if (m_syncingSelectAll || m_data.isEmpty()) {
        return;
    }

    const bool deleted = (state != Qt::Checked);
    for (int i = 0; i < m_data.size(); ++i) {
        if (m_data[i].deleted == deleted) {
            continue;
        }

        m_data[i].deleted = deleted;
        m_dirtyCells.insert(cellKey(m_data[i].row, m_data[i].col));

        auto itemIt = m_itemByCell.find(cellKey(m_data[i].row, m_data[i].col));
        if (itemIt != m_itemByCell.end() && itemIt.value() != nullptr) {
            itemIt.value()->setDeleted(deleted);
        }
    }

    syncPreviewVisibility();
    syncSelectAllState();
}

bool XLSXEditor::eventFilter(QObject* watched, QEvent* event) {
    if (event->type() == QEvent::MouseButtonDblClick) {
        auto* header = qobject_cast<QLabel*>(watched);
        if (header != nullptr && header->property("batchAxis").isValid() &&
            header->property("batchIndex").isValid()) {
            const QString axis = header->property("batchAxis").toString();
            const int index = header->property("batchIndex").toInt();

            bool hasTarget = false;
            bool allKept = true;
            for (const auto& entry : std::as_const(m_data)) {
                if (entry.image.isNull()) {
                    continue;
                }
                if ((axis == "row" && entry.row != index) ||
                    (axis == "col" && entry.col != index)) {
                    continue;
                }
                hasTarget = true;
                if (entry.deleted) {
                    allKept = false;
                    break;
                }
            }

            if (!hasTarget) {
                return true;
            }

            const bool nextDeleted = allKept;
            for (int i = 0; i < m_data.size(); ++i) {
                const auto& entry = m_data[i];
                if (entry.image.isNull()) {
                    continue;
                }
                if ((axis == "row" && entry.row != index) ||
                    (axis == "col" && entry.col != index)) {
                    continue;
                }
                if (m_data[i].deleted == nextDeleted) {
                    continue;
                }

                m_data[i].deleted = nextDeleted;
                m_dirtyCells.insert(cellKey(m_data[i].row, m_data[i].col));

                auto itemIt = m_itemByCell.find(cellKey(m_data[i].row, m_data[i].col));
                if (itemIt != m_itemByCell.end() && itemIt.value() != nullptr) {
                    itemIt.value()->setDeleted(nextDeleted);
                }
            }

            syncPreviewVisibility();
            syncSelectAllState();
            return true;
        }
    }

    if (watched == ui->scrollArea->viewport() && event->type() == QEvent::Wheel) {
        auto* wheelEvent = static_cast<QWheelEvent*>(event);
        if (wheelEvent->modifiers().testFlag(Qt::ControlModifier)) {
            const int delta = wheelEvent->angleDelta().y();
            if (delta == 0) {
                return true;
            }

            const double steps = static_cast<double>(delta) / 120.0;
            const double nextScale = std::clamp(m_itemScale + steps * 0.1, 0.5, 2.5);
            if (std::abs(nextScale - m_itemScale) < 1e-6) {
                return true;
            }

            m_itemScale = nextScale;
            for (auto* item : m_dataItems) {
                if (item) {
                    item->applyScale(m_itemScale);
                }
            }
            updateScrollWidgetSize();
            wheelEvent->accept();
            return true;
        }
    }

    return QWidget::eventFilter(watched, event);
}

bool XLSXEditor::saveData() {
    if (!prepareSaveTargetFile()) {
        return false;
    }

    const bool saved = m_dryRun ? saveFakeDelete() : saveRealDelete();
    if (saved) {
        m_dirtyCells.clear();
    }
    return saved;
}

bool XLSXEditor::prepareSaveTargetFile() {
    if (m_filePath.isEmpty()) {
        return false;
    }

    QFileInfo sourceInfo(m_filePath);
    const QDir sourceDir = sourceInfo.absoluteDir();
    const QString filteredDirPath = sourceDir.filePath("filtered");

    QDir dir;
    if (!dir.mkpath(filteredDirPath)) {
        qWarning() << "Failed to create filtered folder:" << filteredDirPath;
        return false;
    }

    const QString targetPath = QDir(filteredDirPath).filePath(sourceInfo.fileName());
    QFile::remove(targetPath);
    if (!QFile::copy(sourceInfo.absoluteFilePath(), targetPath)) {
        qWarning() << "Failed to copy source xlsx to filtered target:" << targetPath;
        return false;
    }

    if (m_wrapper) {
        m_wrapper->close();
        delete m_wrapper;
        m_wrapper = nullptr;
    }
    if (m_pictureReader.isOpen()) {
        m_pictureReader.close();
    }

    m_saveFilePath = targetPath;
    m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
    if (!m_wrapper->open(m_saveFilePath.toStdString())) {
        delete m_wrapper;
        m_wrapper = nullptr;
        qWarning() << "Failed to open filtered xlsx target:" << m_saveFilePath;
        return false;
    }

    if (!m_pictureReader.open(m_saveFilePath.toStdString())) {
        m_wrapper->close();
        delete m_wrapper;
        m_wrapper = nullptr;
        qWarning() << "Failed to open picture reader on filtered xlsx target:" << m_saveFilePath;
        return false;
    }

    return true;
}

void XLSXEditor::beginSaveProgress(int maximum) {
    if (!m_enableSaveProgress || !ui || !ui->progressBar) {
        return;
    }
    ui->progressBar->setMinimum(0);
    ui->progressBar->setMaximum(std::max(1, maximum));
    ui->progressBar->setValue(0);
    ui->progressBar->setFormat(tx("Saving... %p%"));
    ui->progressBar->setVisible(true);
    QCoreApplication::processEvents();
}

void XLSXEditor::updateSaveProgress(int value) {
    if (!m_enableSaveProgress || !ui || !ui->progressBar || !ui->progressBar->isVisible()) {
        return;
    }
    ui->progressBar->setValue(value);
    QCoreApplication::processEvents();
}

void XLSXEditor::endSaveProgress() {
    if (!m_enableSaveProgress || !ui || !ui->progressBar) {
        return;
    }
    ui->progressBar->setFormat(QStringLiteral("%p%"));
    ui->progressBar->setVisible(false);
    QCoreApplication::processEvents();
}

bool XLSXEditor::saveFakeDelete() {
    const int total = std::max(1, static_cast<int>(m_data.size()) + 1);
    beginSaveProgress(total);

    int progress = 0;
    for (const auto& entry : m_data) {
        QString descCell = numToCol(entry.col) + QString::number(entry.row + 1);
        cc::neolux::utils::MiniXLSX::CellStyle cs;
        cs.backgroundColor = entry.deleted ? "#FF0000" : "";
        m_wrapper->setCellStyle(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(),
                                cs);
        updateSaveProgress(++progress);
    }

    const bool ok = m_wrapper->save();
    updateSaveProgress(total);
    endSaveProgress();
    return ok;
}

bool XLSXEditor::saveRealDelete() {
    const int total = std::max(1, static_cast<int>(m_data.size()) + 8);
    beginSaveProgress(total);
    int progress = 0;

    // 第 1 阶段：先通过 OpenXLSX 写回描述单元格（清空标记删除项）
    // 这样可以确保文本与样式修改由上层接口稳定落盘。
    for (const auto& entry : m_data) {
        QString descCell = numToCol(entry.col) + QString::number(entry.row + 1);
        cc::neolux::utils::MiniXLSX::CellStyle cs;
        cs.backgroundColor = "";
        m_wrapper->setCellStyle(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(),
                                cs);
        if (entry.deleted) {
            m_wrapper->setCellValue(static_cast<unsigned int>(m_sheetIndex), descCell.toStdString(),
                                    "");
        }
        updateSaveProgress(++progress);
    }

    if (!m_wrapper->save()) {
        qWarning() << "fail to save cell modifications";
        endSaveProgress();
        return false;
    }
    updateSaveProgress(++progress);

    // 第 2 阶段：关闭 OpenXLSX 句柄，避免后续 XML 文件改写时的占用冲突。
    m_wrapper->close();
    delete m_wrapper;
    m_wrapper = nullptr;

    m_pictureReader.close();

    // 第 3 阶段：重新解压文件，准备直接修改 drawing 与关系 XML。
    if (!m_pictureReader.open(m_saveFilePath.toStdString())) {
        endSaveProgress();
        return false;
    }
    updateSaveProgress(++progress);

    const std::string tempDir = m_pictureReader.getTempDir();
    if (tempDir.empty()) {
        endSaveProgress();
        return false;
    }

    namespace fs = std::filesystem;
    const fs::path unpackRoot(tempDir);
    const fs::path drawingPath =
        unpackRoot / "xl" / "drawings" / ("drawing" + std::to_string(m_sheetIndex + 1) + ".xml");
    const fs::path drawingRelsPath =
        drawingPath.parent_path() / "_rels" / (drawingPath.filename().string() + ".rels");

    // 没有 drawing 文件说明当前表没有图片对象，直接恢复打开状态即可。
    if (!fs::exists(drawingPath) || !fs::exists(drawingRelsPath)) {
        m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
        if (!m_wrapper->open(m_saveFilePath.toStdString())) {
            endSaveProgress();
            return false;
        }
        updateSaveProgress(total);
        endSaveProgress();
        return true;
    }

    // 收集需要删除的图片坐标（与 drawing 锚点坐标对齐）。
    std::set<std::pair<int, int>> deletedCoords;
    for (const auto& entry : m_data) {
        if (entry.deleted) {
            deletedCoords.insert({entry.row, entry.col});
        }
    }

    pugi::xml_document drawingDoc;
    if (!drawingDoc.load_file(drawingPath.c_str())) {
        endSaveProgress();
        return false;
    }

    pugi::xml_document relsDoc;
    if (!relsDoc.load_file(drawingRelsPath.c_str())) {
        endSaveProgress();
        return false;
    }
    updateSaveProgress(++progress);

    pugi::xml_node wsDr = drawingDoc.child("xdr:wsDr");
    if (!wsDr) {
        wsDr = drawingDoc.first_child();
    }
    if (!wsDr) {
        endSaveProgress();
        return false;
    }

    // 第 4 阶段：删除 drawing.xml 中对应图片锚点，记录 embedId。
    std::unordered_set<std::string> removedEmbedIds;
    for (pugi::xml_node anchor = wsDr.first_child(); anchor;) {
        pugi::xml_node nextAnchor = anchor.next_sibling();
        if (std::string(anchor.name()) != "xdr:twoCellAnchor") {
            anchor = nextAnchor;
            continue;
        }

        pugi::xml_node from = anchor.child("xdr:from");
        const int colNum = from.child("xdr:col").text().as_int(-1) + 1;
        const int rowNum = from.child("xdr:row").text().as_int(-1) + 1;
        if (deletedCoords.find({rowNum, colNum}) == deletedCoords.end()) {
            anchor = nextAnchor;
            continue;
        }

        const std::string embedId = anchor.child("xdr:pic")
                                        .child("xdr:blipFill")
                                        .child("a:blip")
                                        .attribute("r:embed")
                                        .as_string();
        if (!embedId.empty()) {
            removedEmbedIds.insert(embedId);
        }
        wsDr.remove_child(anchor);
        anchor = nextAnchor;
    }

    if (!drawingDoc.save_file(drawingPath.c_str(), PUGIXML_TEXT("  "))) {
        endSaveProgress();
        return false;
    }
    updateSaveProgress(++progress);

    pugi::xml_node relRoot = relsDoc.child("Relationships");
    if (!relRoot) {
        relRoot = relsDoc.first_child();
    }
    if (!relRoot) {
        endSaveProgress();
        return false;
    }

    // 第 5 阶段：在 drawing rels 中删除 embedId 对应关系并记录图片目标路径。
    std::unordered_set<std::string> removedTargets;
    for (pugi::xml_node rel = relRoot.child("Relationship"); rel;) {
        pugi::xml_node nextRel = rel.next_sibling("Relationship");
        const std::string id = rel.attribute("Id").as_string();
        if (removedEmbedIds.find(id) != removedEmbedIds.end()) {
            const std::string target = rel.attribute("Target").as_string();
            if (!target.empty()) {
                removedTargets.insert(target);
            }
            relRoot.remove_child(rel);
        }
        rel = nextRel;
    }

    if (!relsDoc.save_file(drawingRelsPath.c_str(), PUGIXML_TEXT("  "))) {
        endSaveProgress();
        return false;
    }
    updateSaveProgress(++progress);

    // 第 6 阶段：删除已不再被任何关系引用的 media 图片文件。
    std::unordered_set<std::string> activeImageTargets;
    for (pugi::xml_node rel = relRoot.child("Relationship"); rel;
         rel = rel.next_sibling("Relationship")) {
        const std::string type = rel.attribute("Type").as_string();
        if (type.find("/image") != std::string::npos) {
            activeImageTargets.insert(rel.attribute("Target").as_string());
        }
    }

    for (const auto& target : removedTargets) {
        if (activeImageTargets.find(target) != activeImageTargets.end()) {
            continue;
        }
        const fs::path imagePath =
            (drawingPath.parent_path() / fs::path(target)).lexically_normal();
        if (fs::exists(imagePath)) {
            std::error_code ec;
            fs::remove(imagePath, ec);
        }
    }

    // 第 7 阶段：将修改后的临时目录重新打包为 xlsx。
    // 注意：打包前不能 close pictureReader，否则临时目录会被清理。
    if (!cc::neolux::utils::KFZippa::zip(unpackRoot.string(), m_saveFilePath.toStdString())) {
        qWarning() << "Failed to repack XLSX file";
        m_pictureReader.close();
        // Try to reopen anyway
        m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
        if (!m_wrapper->open(m_saveFilePath.toStdString())) {
            endSaveProgress();
            return false;
        }
        if (!m_pictureReader.open(m_saveFilePath.toStdString())) {
            endSaveProgress();
            return false;
        }
        endSaveProgress();
        return false;
    }
    updateSaveProgress(++progress);

    // 打包完成后再清理临时目录。
    m_pictureReader.close();

    // 第 8 阶段：重新打开读取器，保持编辑器处于可继续操作状态。
    m_wrapper = new cc::neolux::utils::MiniXLSX::OpenXLSXWrapper();
    if (!m_wrapper->open(m_saveFilePath.toStdString())) {
        qWarning() << "Failed to reopen OpenXLSX after repacking";
        endSaveProgress();
        return false;
    }

    if (!m_pictureReader.open(m_saveFilePath.toStdString())) {
        qWarning() << "Failed to reopen XLPictureReader after repacking";
        endSaveProgress();
        return false;
    }

    updateSaveProgress(total);
    endSaveProgress();

    return true;
}

void XLSXEditor::showImageDialog(int row, int col) {
    for (const auto& entry : m_data) {
        if (entry.row == row && entry.col == col && !entry.image.isNull()) {
            QDialog dialog(this);
            dialog.setWindowTitle(tx("Image Preview"));
            QLabel* label = new QLabel(&dialog);
            label->setPixmap(QPixmap::fromImage(entry.image));
            QVBoxLayout* layout = new QVBoxLayout(&dialog);
            layout->addWidget(label);
            dialog.setLayout(layout);
            dialog.exec();
            break;
        }
    }
}

void XLSXEditor::updateScrollWidgetSize() {
    int startRow, startCol, endRow, endCol;
    parseRange(m_range, startRow, startCol, endRow, endCol);

    QSet<int> rowSet;
    QSet<int> colSet;
    for (const auto& entry : std::as_const(m_data)) {
        if (!entry.image.isNull()) {
            rowSet.insert(entry.row);
            colSet.insert(entry.col);
        }
    }
    const int rowCount = rowSet.size();
    const int colCount = colSet.size();
    const int itemW = static_cast<int>(std::round(kBaseItemWidth * m_itemScale));
    const int itemH = static_cast<int>(std::round(kBaseItemHeight * m_itemScale));
    const int headerW = static_cast<int>(std::round(kBaseHeaderColWidth * m_itemScale));
    const int headerH = static_cast<int>(std::round(kBaseHeaderRowHeight * m_itemScale));

    const int contentW =
        colCount > 0 ? headerW + colCount * itemW + colCount * kGridSpacing : headerW;
    const int contentH =
        rowCount > 0 ? headerH + rowCount * itemH + rowCount * kGridSpacing : headerH;

    ui->scrollWidget->resize(contentW, contentH);
    ui->scrollWidget->setMinimumSize(contentW, contentH);
}

void XLSXEditor::syncPreviewButtonText() {
    if (!ui || !ui->btnPreview) {
        return;
    }
    ui->btnPreview->setText(m_previewOnly ? tx("Show All") : tx("Preview"));
}

void XLSXEditor::syncPreviewVisibility() {
    for (int i = 0; i < m_data.size(); ++i) {
        const auto& entry = m_data[i];
        auto itemIt = m_itemByCell.find(cellKey(entry.row, entry.col));
        if (itemIt == m_itemByCell.end() || itemIt.value() == nullptr) {
            continue;
        }
        const bool visible = !m_previewOnly || !entry.deleted;
        itemIt.value()->setVisible(visible);
    }
}

void XLSXEditor::syncSelectAllState() {
    if (!ui || !ui->chkSelectAll) {
        return;
    }

    m_syncingSelectAll = true;
    if (m_data.isEmpty()) {
        ui->chkSelectAll->setCheckState(Qt::Unchecked);
        m_syncingSelectAll = false;
        return;
    }

    int deletedCount = 0;
    for (const auto& entry : std::as_const(m_data)) {
        if (entry.deleted) {
            ++deletedCount;
        }
    }

    const bool allKept = (deletedCount == 0);
    ui->chkSelectAll->setCheckState(allKept ? Qt::Checked : Qt::Unchecked);
    m_syncingSelectAll = false;
}
