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

class DataItem : public QWidget {
    Q_OBJECT

public:
    explicit DataItem(QWidget* parent = nullptr);
    ~DataItem();

    void setImage(const QImage& image);
    void setDescription(const QString& desc);
    QString getDescription() const;
    void setDeleted(bool deleted);
    bool isDeleted() const;
    void setPosition(int row, int col);
    void setRowCol(int row, int col);
    int getRow() const;
    int getCol() const;
    void applyScale(double scale);

signals:
    void imageClicked(int row, int col);
    void deleteToggled(bool deleted);

protected:
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
