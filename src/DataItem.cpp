#include "cc/neolux/fem/xlsxeditor/DataItem.hpp"

#include <QCoreApplication>
#include <QEvent>
#include <QLabel>
#include <QPixmap>
#include <QSizePolicy>
#include <algorithm>
#include <cmath>

#include "ui_DataItem.h"

namespace {
constexpr int kBaseItemWidth = 70;
constexpr int kBaseItemHeight = 90;
constexpr int kBaseContentSize = 68;
constexpr int kBaseIconSize = 66;
constexpr double kInnerGapPercent = 0.3;

QString tx(const char* sourceText) {
    return QCoreApplication::translate("DataItem", sourceText);
}
}  // namespace

namespace cc::neolux::fem::xlsxeditor {

DataItem::DataItem(QWidget* parent)
    : QWidget(parent), ui(new Ui::DataItem), m_deleted(true), m_row(-1), m_col(-1) {
    ui->setupUi(this);
    setAttribute(Qt::WA_StyledBackground, true);
    setStyleSheet("#DataItem { border: 1px solid #606060; }");
    QSizePolicy policy(QSizePolicy::Fixed, QSizePolicy::Fixed);
    policy.setRetainSizeWhenHidden(true);
    setSizePolicy(policy);
    ui->gridLayout->setContentsMargins(0, 0, 0, 0);
    ui->gridLayout->setSpacing(0);
    ui->btnImage->setFlat(true);
    ui->btnImage->setStyleSheet("QPushButton { border: 0; padding: 0; margin: 0; }");
    // 改为通过悬停预览：安装事件过滤器以捕获 Enter/Leave
    ui->btnImage->installEventFilter(this);
    ui->lnData->setReadOnly(true);  // 数据只读，防止误修改
    ui->lnData->setToolTip(tx("Double-click to keep/remove"));
    ui->lnData->installEventFilter(this);
    const qreal basePointSize =
        ui->lnData->font().pointSizeF() > 0 ? ui->lnData->font().pointSizeF() : 9.0;
    ui->lnData->setProperty("basePointSizeF", basePointSize);
    ui->lnData->setProperty("baseHeight", ui->lnData->sizeHint().height());
    applyScale(1.0);
}

DataItem::~DataItem() {
    delete ui;
}

void DataItem::setImage(const QImage& image) {
    m_image = image;
    if (m_image.isNull()) {
        ui->btnImage->setIcon(QIcon());
        return;
    }

    const QSize iconSize = ui->btnImage->iconSize();
    const int side = std::max(iconSize.width(), iconSize.height());
    const int renderSide = side > 0 ? side : kBaseIconSize;
    const QPixmap pixmap = QPixmap::fromImage(
        m_image.scaled(renderSide, renderSide, Qt::KeepAspectRatio, Qt::SmoothTransformation));
    ui->btnImage->setIcon(QIcon(pixmap));
}

void DataItem::setDescription(const QString& desc) {
    ui->lnData->setText(desc);
}

QString DataItem::getDescription() const {
    return ui->lnData->text();
}

void DataItem::setDeleted(bool deleted) {
    m_deleted = deleted;
    if (deleted) {
        ui->lnData->setStyleSheet("background-color: #ff4d4f;");
    } else {
        ui->lnData->setStyleSheet("");
    }
}

bool DataItem::isDeleted() const {
    return m_deleted;
}

void DataItem::setRowCol(int row, int col) {
    m_row = row;
    m_col = col;
}

int DataItem::getRow() const {
    return m_row;
}

int DataItem::getCol() const {
    return m_col;
}

void DataItem::applyScale(double scale) {
    const double clamped = std::clamp(scale, 0.5, 2.5);
    const double gapRatio = std::clamp(kInnerGapPercent, 0.0, 20.0) / 100.0;
    const int itemW = static_cast<int>(std::round(kBaseItemWidth * clamped));
    const int itemH = static_cast<int>(std::round(kBaseItemHeight * clamped));
    const int contentSize = static_cast<int>(std::round(kBaseContentSize * clamped));
    const int iconSize = static_cast<int>(std::round(kBaseIconSize * clamped));
    const int innerGap = std::max(1, static_cast<int>(std::round(itemW * gapRatio)));
    const qreal basePointSize = ui->lnData->property("basePointSizeF").toReal();

    setFixedSize(itemW, itemH);
    ui->gridLayout->setContentsMargins(innerGap, innerGap, innerGap, innerGap);
    ui->gridLayout->setSpacing(innerGap);
    ui->btnImage->setFixedSize(contentSize, contentSize);
    ui->btnImage->setIconSize(QSize(iconSize, iconSize));
    const int textWidth = itemW - 2 * innerGap;
    const int textHeight = itemH - contentSize - 3 * innerGap;
    ui->lnData->setFixedWidth(textWidth);
    ui->lnData->setFixedHeight(std::max(14, textHeight));

    QFont textFont = ui->lnData->font();
    textFont.setPointSizeF(std::max<qreal>(6.0, basePointSize * clamped));
    ui->lnData->setFont(textFont);

    if (!m_image.isNull()) {
        const QPixmap pixmap = QPixmap::fromImage(
            m_image.scaled(iconSize, iconSize, Qt::KeepAspectRatio, Qt::SmoothTransformation));
        ui->btnImage->setIcon(QIcon(pixmap));
    }
}

bool DataItem::eventFilter(QObject* watched, QEvent* event) {
    if (watched == ui->lnData && event->type() == QEvent::MouseButtonDblClick) {
        setDeleted(!m_deleted);
        emit deleteToggled(m_deleted);
        return true;
    }

    if (watched == ui->btnImage) {
        // 鼠标进入图片按钮：通知父组件显示悬停预览
        if (event->type() == QEvent::Enter) {
            emit imageEntered(m_row, m_col);
            return true;
        }
        // 鼠标离开图片按钮：通知父组件隐藏（或延迟隐藏）悬停预览
        else if (event->type() == QEvent::Leave) {
            emit imageLeft(m_row, m_col);
            return true;
        }
    }
    return QWidget::eventFilter(watched, event);
}

/**
 * @brief 返回当前项持有的图片（原始 QImage）。
 */
QImage DataItem::getImage() const {
    return m_image;
}

/**
 * @brief 返回图片按钮在屏幕全局坐标系中的左上角位置。
 *
 * 该接口用于在主界面中定位悬停预览窗格时将 item 的本地位置转换为全局坐标。
 */
QPoint DataItem::imageWidgetGlobalPos() const {
    if (!ui || !ui->btnImage) {
        return QPoint();
    }
    return ui->btnImage->mapToGlobal(QPoint(0, 0));
}

}  // namespace cc::neolux::fem::xlsxeditor
