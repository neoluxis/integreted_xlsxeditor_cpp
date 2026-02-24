#pragma once

#include <QHash>
#include <QImage>
#include <QProgressBar>
#include <QSet>
#include <QString>
#include <QVector>
#include <QWidget>
#include <cc/neolux/utils/MiniXLSX/OpenXLSXWrapper.hpp>
#include <cc/neolux/utils/MiniXLSX/XLPictureReader.hpp>

namespace Ui {
class XLSXEditor;
}

namespace cc {
namespace neolux {
namespace fem {
namespace xlsxeditor {

/**
 * @brief 单个数据项（图片 + 描述）的内存表示。
 *
 * row/col 使用工作表中的 1-based 行列坐标。
 */
struct DataEntry {
    int row, col;
    QImage image;
    QString desc;
    bool deleted;
};

class DataItem;

/**
 * @brief XLSX 编辑器主界面组件。
 *
 * 负责加载指定工作表和范围内的数据项，并支持：
 * - 预览已保留项
 * - 标记删除/取消删除
 * - 假删除保存（仅标红）与真删除保存（删除图片与描述）
 */
class XLSXEditor : public QWidget {
    Q_OBJECT

public:
    /**
     * @brief 构造编辑器。
     * @param parent 父级 QWidget。
     * @param dry_run 是否使用假删除模式（true: 仅标记；false: 真删除）。
     */
    explicit XLSXEditor(QWidget* parent = nullptr, bool dry_run = true);

    /**
     * @brief 析构函数，释放 UI、XLSX 包装器与图片读取器资源。
     */
    ~XLSXEditor();

    /**
     * @brief 加载 XLSX 文件指定表和范围的数据。
     * @param filePath XLSX 文件路径。
     * @param sheetName 目标工作表名称。
     * @param range 读取范围，支持 A:C,1:10 或 A1:C10 格式。
     */
    void loadXLSX(const QString& filePath, const QString& sheetName, const QString& range);

    /**
     * @brief 设置删除模式。
     * @param dry_run true 为假删除（标红），false 为真删除（删除图片与描述）。
     */
    void setDryRun(bool dry_run);

    /**
     * @brief 获取当前删除模式。
     * @return true 表示假删除模式，false 表示真删除模式。
     */
    bool isDryRun() const;

private slots:
    /** @brief 处理“保存”按钮点击事件。 */
    void on_btnSave_clicked();

    /** @brief 处理“预览/显示全部”按钮点击事件。 */
    void on_btnPreview_clicked();

    /**
     * @brief 处理“全选”复选框状态变化。
     * @param state Qt 复选框状态值。
     */
    void on_chkSelectAll_stateChanged(int state);

protected:
    /**
     * @brief 事件过滤器，用于处理滚轮缩放等交互。
     * @param watched 事件源对象。
     * @param event 事件对象。
     * @return true 表示事件已处理，false 表示继续默认分发。
     */
    bool eventFilter(QObject* watched, QEvent* event) override;

private:
    Ui::XLSXEditor* ui;
    QString m_filePath;
    QString m_sheetName;
    QString m_range;
    QVector<DataEntry> m_data;  // 图片与描述及其位置
    QVector<DataItem*> m_dataItems;
    QHash<QString, int> m_indexByCell;
    QHash<QString, DataItem*> m_itemByCell;
    QSet<QString> m_dirtyCells;
    cc::neolux::utils::MiniXLSX::OpenXLSXWrapper* m_wrapper;
    cc::neolux::utils::MiniXLSX::XLPictureReader m_pictureReader;
    int m_sheetIndex;

    /**
     * @brief 解析范围字符串为起止行列。
     * @param range 范围文本，支持 A:C,1:10 或 A1:C10。
     * @param startRow 起始行（输出）。
     * @param startCol 起始列（输出）。
     * @param endRow 结束行（输出）。
     * @param endCol 结束列（输出）。
     */
    void parseRange(const QString& range, int& startRow, int& startCol, int& endRow, int& endCol);

    /** @brief 加载当前范围内的数据（无进度条版本）。 */
    void loadData();

    /**
     * @brief 加载当前范围内的数据，并更新进度条。
     * @param progressBar 进度条对象引用。
     */
    void loadData(QProgressBar& progressBar);

    /**
     * @brief 将加载到的数据渲染到界面网格。
     * @param previewOnly true 时仅显示未删除项。
     */
    void displayData(bool previewOnly = false);

    /**
     * @brief 保存当前编辑结果。
     * @return 保存成功返回 true，否则返回 false。
     */
    bool saveData();

    /**
     * @brief 假删除保存：仅写回描述背景色标记。
     * @return 保存成功返回 true，否则返回 false。
     */
    bool saveFakeDelete();

    /**
     * @brief 真删除保存：删除图片锚点、关系与描述内容。
     * @return 保存成功返回 true，否则返回 false。
     */
    bool saveRealDelete();

    /**
     * @brief 将列字母转换为列号。
     * @param col 列字母（如 A、AB）。
     * @return 1-based 列号。
     */
    int colToNum(const QString& col);

    /**
     * @brief 将列号转换为列字母。
     * @param num 1-based 列号。
     * @return 列字母（如 A、AB）。
     */
    QString numToCol(int num);

    /**
     * @brief 生成单元格键值（row:col）。
     * @param row 1-based 行号。
     * @param col 1-based 列号。
     * @return 用于哈希索引的字符串键。
     */
    QString cellKey(int row, int col) const;

    /**
     * @brief 弹出图片预览对话框。
     * @param row 图片所在行。
     * @param col 图片所在列。
     */
    void showImageDialog(int row, int col);

    /** @brief 清理并销毁当前界面中的 DataItem 组件。 */
    void clearDataItems();

    /** @brief 重置编辑器内部状态。 */
    void resetState();

    /** @brief 根据当前缩放和范围更新滚动区内容尺寸。 */
    void updateScrollWidgetSize();

    /** @brief 同步预览按钮文本（Preview/Show All）。 */
    void syncPreviewButtonText();

    /** @brief 根据预览状态刷新数据项可见性。 */
    void syncPreviewVisibility();

    /** @brief 根据数据删除状态同步“全选”复选框状态。 */
    void syncSelectAllState();

    /**
     * @brief 开始保存进度显示。
     * @param maximum 进度最大值。
     */
    void beginSaveProgress(int maximum);

    /**
     * @brief 更新保存进度值。
     * @param value 当前进度值。
     */
    void updateSaveProgress(int value);

    /** @brief 结束保存进度显示。 */
    void endSaveProgress();

    /**
     * @brief 是否启用保存进度条。
     *
     * 直接在代码中改此变量即可全局开关保存进度显示。
     */
    bool m_enableSaveProgress;

    bool m_dryRun;
    bool m_previewOnly;
    double m_itemScale;
    bool m_syncingSelectAll;
};

}  // namespace xlsxeditor
}  // namespace fem
}  // namespace neolux
}  // namespace cc
