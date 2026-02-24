#pragma once

#include <QImage>
#include <QString>
#include <QWidget>

namespace Ui {
class DataItem;
}

namespace cc {
namespace neolux {
namespace fem {
namespace xlsxeditor {

/**
 * @brief 单个图片-描述数据展示组件。
 *
 * 该组件用于在网格中展示一组数据：
 * - 图片按钮（点击可预览）
 * - 描述输入框（双击切换删除状态）
 */
class DataItem : public QWidget {
    Q_OBJECT

public:
    /**
     * @brief 构造函数。
     * @param parent 父级 QWidget。
     */
    explicit DataItem(QWidget* parent = nullptr);

    /** @brief 析构函数。 */
    ~DataItem();

    /**
     * @brief 设置展示图片。
     * @param image 图片对象。
     */
    void setImage(const QImage& image);

    /**
     * @brief 设置描述文本。
     * @param desc 描述内容。
     */
    void setDescription(const QString& desc);

    /**
     * @brief 获取当前描述文本。
     * @return 描述内容。
     */
    QString getDescription() const;

    /**
     * @brief 设置删除状态。
     * @param deleted true 表示标记删除，false 表示保留。
     */
    void setDeleted(bool deleted);

    /**
     * @brief 获取删除状态。
     * @return true 表示已标记删除。
     */
    bool isDeleted() const;

    /**
     * @brief 设置行列坐标（兼容接口）。
     * @param row 1-based 行号。
     * @param col 1-based 列号。
     */
    void setPosition(int row, int col);

    /**
     * @brief 设置行列坐标。
     * @param row 1-based 行号。
     * @param col 1-based 列号。
     */
    void setRowCol(int row, int col);

    /**
     * @brief 获取行号。
     * @return 1-based 行号。
     */
    int getRow() const;

    /**
     * @brief 获取列号。
     * @return 1-based 列号。
     */
    int getCol() const;

    /**
     * @brief 应用缩放比例，调整组件尺寸和字体。
     * @param scale 缩放因子。
     */
    void applyScale(double scale);

signals:
    /**
     * @brief 图片点击信号。
     * @param row 当前项行号。
     * @param col 当前项列号。
     */
    void imageClicked(int row, int col);

    /**
     * @brief 删除状态切换信号。
     * @param deleted 新的删除状态。
     */
    void deleteToggled(bool deleted);

protected:
    /**
     * @brief 事件过滤器，用于处理描述框双击切换删除状态。
     * @param watched 事件源对象。
     * @param event 事件对象。
     * @return true 表示事件已处理，false 表示继续默认分发。
     */
    bool eventFilter(QObject* watched, QEvent* event) override;

private:
    Ui::DataItem* ui;
    bool m_deleted;
    int m_row, m_col;
    QImage m_image;
};

}  // namespace xlsxeditor
}  // namespace fem
}  // namespace neolux
}  // namespace cc
