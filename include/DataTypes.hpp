#pragma once

#include <QImage>
#include <QString>

struct DataEntry {
    int row, col;
    QImage image;
    QString desc;
    bool deleted;
};