#ifndef EXCELPROCESS_H
#define EXCELPROCESS_H

#include <QtCore/qglobal.h>
#include "ISheetT.h"
#include "libxl.h"

#include <QMap>
#include <QDebug>
#include <QObject>
#include <conio.h>
#include <iostream>
#include <windows.h>
#include <QStringList>

#include "include.h"


using namespace libxl;

#ifdef EXCELPROCESS_LIB
# define EXCELPROCESS_EXPORT Q_DECL_EXPORT
#else
# define EXCELPROCESS_EXPORT Q_DECL_IMPORT
#endif

class EXCELPROCESS_EXPORT ExcelProcess : public QObject
{
public:
	ExcelProcess(QObject *parent);
	~ExcelProcess();

	////��ȡExcel�ļ�����
	//bool FileData(const QString strFileName, QStringList& arrField, QList<QStringList>& arrFileData, QString strFlag = QString::fromLocal8Bit(""));

	//��ȡExcel�ļ�����
	bool FileData(const QString strFileName, QStringList& arrField, QList<QStringList>& arrFileData, QString strFlag = QString::fromLocal8Bit(""));

	//�������ݵ��ļ�
	bool SaveToFile(const QMap<QString, QStringList>& arrData, QString strFileName);
	bool SaveToFile(QList<ProjInfo> arrProInfo, const QString strFileName);
	bool SaveToFile(ProjInfo info, QStringList arrField, QList<QStringList> arrFileData, const QString strFileName, const QList<int> arrNumIndex, QList<double> arrZJ);

	//������ѯ��Ŀ������Ϣ
	bool SaveToFile(QStringList arrField, QList<QStringList> arrFileData, Book* book, Sheet* sheet, const QList<int> arrNumIndex, QList<double> arrZJ, int iSize);
	//bool SaveToFile(QStringList arrField, QList<QStringList> arrFileData, const QString strFileName, const QList<int> arrNumIndex, QList<double> arrZJ, int iSize);
	bool SaveToFile(QList<QStringList> arr, const QString strFileName, QList<int> arrNumIndex, int iSheetIndex);

	//����ERP���δ���������
	bool SaveToFile(QList<QStringList> arrData, QString strFileName);

	//������ѯ��Ŀ����������������
	bool SaveToFile(QString strFileName, QList<QStringList> arrData1, QList<QStringList> arrData2);

	//������е����ݽ��з���������ͬ�ط�����ɫ������ʾ
	void ExcelDataCompare(const QString strFileName);

private:

private:
	QMap<QString, QStringList> m_arrFileData;
	QList<QString> m_WZFlags;
	QList<QString> m_WZBFlags;

	//����Ƿ��д����ַ���������
	QList<QString> m_XMCheck;
	QList<QString> m_WZCheck;
	QList<QString> m_SGCheck;
	QList<QString> m_JSCheck;
	QList<QString> m_WZBCheck;
};

#endif // EXCELPROCESS_H
