#pragma once
#include <QList>
#include <QTableWidget>

//��Ŀ��Ϣ ���ע�ͣ���ν����ͻ��

struct ProjInfo
{
	char ID[50];		//��Ŀ����
	char Name[100];		//��Ŀ����
	//char state[20];		//�������
	int	 iYear;			//���
	char PC[20];		//����
	char type[20];		//�ʽ�����
	char CoLtd[50];		//ʩ����λ
	double fTZ;			//��Ŀ��Ͷ��
	char Time[20];		//��Ŀ����ʱ��
};



static void MapDataToTable(const QStringList arrField, const QList<QStringList>& arrData, QTableWidget* pTable, const QStringList arrInt, const QStringList arrFloat)
{
	if (pTable && arrField.size() > 0)
	{
		//pTable->setRowCount(0);
		//pTable->setColumnCount(0);

		int iMaxRowNum = arrData[0].size();
		for(int i = 0; i < arrData.size(); i++)
		{
			if(arrData[i].size() > iMaxRowNum)
			{
				iMaxRowNum = arrData[i].size();
			}
		}
		pTable->setRowCount(iMaxRowNum);
		pTable->setColumnCount(arrField.size());

		for (int i = 0; i < arrField.size(); i++)
		{
			//��ӱ�ͷ
			QTableWidgetItem* pHeaderItem = new QTableWidgetItem(arrField[i]);
			pHeaderItem->setTextAlignment(Qt::AlignCenter);
			pTable ->setHorizontalHeaderItem(i, pHeaderItem);

			//���������
			QStringList arrValue = arrData[i];
			for (int j=0; j < arrValue.size(); j++)
			{
				QTableWidgetItem* pItem = new QTableWidgetItem(arrValue[j]);
				pTable->setItem(j,i, pItem);
			}
		}
	}
}

static void MapDataToValueTable(const QStringList arrField, const QMap<QString,bool> arrFielvVisable, const QList<QStringList>& arrData, QTableWidget* pTable, const QStringList arrInt, const QStringList arrFloat, QList<double>& arrZJ)
{
	if (pTable && arrField.size() > 0)
	{
		pTable->setRowCount(0);
		pTable->setColumnCount(0);

		int iMaxRowNum = 0;
		for(int i = 0; i < arrData.size(); i++)
		{
			if(arrData[i].size() > iMaxRowNum)
			{
				iMaxRowNum = arrData[i].size();
			}
		}
		pTable->setRowCount(iMaxRowNum+3);
		pTable->setColumnCount(arrField.size());

		//���ñ���ͷ
		for(int j = 0; j < 3; j++)
		{
			for(int i = 0; i < arrField.size(); i++)
			{
				QTableWidgetItem* pItem = new QTableWidgetItem(QString::fromLocal8Bit(""));
				pItem->setTextAlignment(Qt::AlignCenter);
				QFont fnt = pItem->font();
				fnt.setBold(true);
				pItem->setFont(fnt);
				pTable->setItem(j,i,pItem);
			}
		}
		pTable->setSpan(0,0,3,1);
		pTable->setSpan(0,1,3,1);
		pTable->setSpan(0,2,3,1);
		pTable->setSpan(0,3,3,1);
		pTable->setSpan(0,4,3,1);
		pTable->setSpan(0,5,3,1);
		pTable->setSpan(0,6,3,1);
		pTable->item(0,0)->setText(QString::fromLocal8Bit("���ʱ���&��չ��"));
		pTable->item(0,1)->setText(QString::fromLocal8Bit("�����"));
		pTable->item(0,2)->setText(QString::fromLocal8Bit("�з���"));
		pTable->item(0,3)->setText(QString::fromLocal8Bit("С����"));
		pTable->item(0,4)->setText(QString::fromLocal8Bit("��������"));
		pTable->item(0,5)->setText(QString::fromLocal8Bit("��չ����"));
		pTable->item(0,6)->setText(QString::fromLocal8Bit("�����淶ID"));

		//����Э���
		pTable->setSpan(0,7,1,2);
		pTable->setSpan(1,7,1,2);
		pTable->item(0,7)->setText(QString::fromLocal8Bit("Э�����ʱ�"));
		pTable->item(2,7)->setText(QString::fromLocal8Bit("��λ"));
		pTable->item(2,8)->setText(QString::fromLocal8Bit("����"));
		
		//�������
		pTable->setSpan(0,9,1,3);
		pTable->setSpan(1,9,1,3);
		pTable->item(0,9)->setText(QString::fromLocal8Bit("�������"));
		pTable->item(2,9)->setText(QString::fromLocal8Bit("��λ"));
		pTable->item(2,10)->setText(QString::fromLocal8Bit("����"));
		pTable->item(2,11)->setText(QString::fromLocal8Bit("����"));

		//ʩ�����
		pTable->setSpan(0,12,1,1);
		pTable->setSpan(1,12,1,1);
		pTable->item(0,12)->setText(QString::fromLocal8Bit("ʩ�����"));
		pTable->item(2,12)->setText(QString::fromLocal8Bit("����"));

		//�������
		pTable->setSpan(0,13,1,2);
		pTable->setSpan(1,13,1,2);
		pTable->item(0,13)->setText(QString::fromLocal8Bit("�������"));
		pTable->item(2,13)->setText(QString::fromLocal8Bit("��λ"));
		pTable->item(2,14)->setText(QString::fromLocal8Bit("����"));

		//ERP
		pTable->setSpan(0,15,1,3);
		pTable->setSpan(1,15,1,3);
		pTable->item(0,15)->setText("ERP");
		pTable->item(2,15)->setText("��λ");
		pTable->item(2,16)->setText("����");
		pTable->item(2,17)->setText("����");

		for (int i = 0; i < arrField.size(); i++)
		{
			//���������
			QStringList arrValue = arrData[i];
			for (int j=0; j < arrValue.size(); j++)
			{
				QTableWidgetItem* pItem = new QTableWidgetItem(arrValue[j]);
				pTable->setItem(j+3,i, pItem);
			}
			
			if(!arrFielvVisable[arrField.at(i)])
			{
				pTable->hideColumn(i);
			}
		}

		/*
		//ͳ�Ƹ�����ܼ�
		double fWZBZJ= 0;	//���ʱ��ܼ�
		double fWZQCZJ= 0; //��������ܼ�
		double fSGZJ= 0;	//ʩ������ܼ�
		double fJSZJ= 0;	//��������ܼ�
		double fERPZJ=0;	//ERP�ܼ�
		int iRow = arrData[0].size();
		for (int i = 0; i < iRow; i++)
		{
			//���������
			for (int j=0; j < arrField.size(); j++)
			{
				double fWZBDJ = arrData[8] .at(j).toDouble();	//���ʱ���
				double fQCDJ  = arrData[11].at(j).toDouble();	//������ᵥ��
				double fQCSL  = arrData[10].at(j).toDouble();	//�����������
				double fSGSL  = arrData[12].at(j).toDouble();	//ʩ���������
				double fJSSL  = arrData[14].at(j).toDouble();	//�����������
				double fERPDJ = arrData[16].at(j).toDouble();	//ERP����
				double fERPSL = arrData[17].at(j).toDouble();	//ERP����

				fWZBZJ		+= fWZBDJ*fQCSL;
				fWZQCZJ		+= fQCDJ*fQCSL;
				fSGZJ		+= fQCDJ*fQCSL;
				fJSZJ		+= fQCDJ*fQCSL;
				fERPZJ		+= fERPDJ*fERPSL;
			}
		}

		arrZJ << fWZBZJ <<fWZQCZJ << fSGZJ << fJSZJ << fERPZJ;
		*/

		pTable->setItem(1,7,	new QTableWidgetItem("�ܼ�: "+QString::number(arrZJ[0],'f',2)));
		pTable->setItem(1,9,	new QTableWidgetItem("�ܼ�: "+QString::number(arrZJ[1],'f',2)));
		pTable->setItem(1,12,	new QTableWidgetItem("�ܼ�: "+QString::number(arrZJ[2],'f',2)));
		pTable->setItem(1,13,	new QTableWidgetItem("�ܼ�: "+QString::number(arrZJ[3],'f',2)));
		//pTable->setItem(1,12, new QTableWidgetItem("�ܼ�"));
		//pTable->setItem(1,13, new QTableWidgetItem("�ܼ�"));
		pTable->setItem(1,15,	new QTableWidgetItem("�ܼ�: "+QString::number(arrZJ[4],'f',2)));
	}
}

static void TableData2Map(QTableWidget* pTable, QStringList& arrOldField, QList<QStringList>& arrOldData, QStringList& arrNewField, QList<QStringList>& arrNewData)
{
	int iRowCnt = pTable->rowCount();
	for(int i = 0; i < arrOldField.size(); i++)
	{
		arrNewField << pTable->horizontalHeaderItem(i)->text();

		QString strData;
		for(int j = 0; j < iRowCnt; j++)
		{
			strData += pTable->item(j,i)->text();
			if(j != iRowCnt-1)
			{
				strData += QString::fromLocal8Bit(";");
			}
		}
		QStringList arrData = strData.split(";");
		arrNewData << arrData;
	}
}

static void ValueTableData2Map(QTableWidget* pTable, QStringList& arrOldField, QList<QStringList>& arrOldData, QStringList& arrNewField, QList<QStringList>& arrNewData)
{
	int iRowCnt = pTable->rowCount();
	for(int i = 0; i < arrOldField.size(); i++)
	{
		QString strData;
		for(int j = 2; j < iRowCnt; j++)
		{
			strData += pTable->item(j,i)->text();
			if(j != iRowCnt-1)
			{
				strData += QString::fromLocal8Bit(";");
			}
		}
		QStringList arrData = strData.split(";");
		arrNewData << arrData;
	}
}

#include <QMessageBox>
//iType 0��Ŀ��1���ʡ�2ʩ����3���㡢4���ʱ�
static bool check(QString iType, QStringList arrField)
{
	QStringList arr;
	if (iType == "@XM")
	{
		arr << QString::fromLocal8Bit("���") << QString::fromLocal8Bit("����") << QString::fromLocal8Bit("���������")<<QString::fromLocal8Bit("��Ƶ�λ")<<QString::fromLocal8Bit("��Ŀ����")<<QString::fromLocal8Bit("�ʽ�����")
			<< QString::fromLocal8Bit("��Ŀ��Ͷ��");
	}
	else if (iType == "@WZ")
	{
		arr <<"������" << "�������"<< "���ʱ���" << "��չ��" <<"��������"<<"��չ������"<<"��������"<<"��Ƶ�λ"<<"�������"
			<<"�ϱ���λ"<<"�ϱ�����"<<"����"<<"�ϼۣ�Ԫ��"<<"����\n��T��"<<"����\n��T��";
	}
	else if(iType == "@SG")
	{
		arr << QString::fromLocal8Bit("���ʱ���") << QString::fromLocal8Bit("��չ��")<<QString::fromLocal8Bit("ʩ���ϱ�����")<<QString::fromLocal8Bit("ʩ���ϱ�����");
	}
	else if (iType == "@JS")
	{
		arr << QString::fromLocal8Bit("���ʱ���") << QString::fromLocal8Bit("��չ��")<<QString::fromLocal8Bit("��������")<<QString::fromLocal8Bit("�����ܼ�");
	}
	else if(iType == "@WZB")
	{
		arr << QString::fromLocal8Bit("����&��չ") << QString::fromLocal8Bit("�����")<<QString::fromLocal8Bit("�з���")<<QString::fromLocal8Bit("С����")
			<<QString::fromLocal8Bit("�����淶ID")<<QString::fromLocal8Bit("��λ")<<QString::fromLocal8Bit("�۸�");
	}

	bool bRet = true;
	QString strMsg = QString::fromLocal8Bit("ȱ��");
	for (int i = 0; i<arr.size(); i++)
	{
		if(!arrField.contains(arr[i]))
		{
			bRet = false;
			strMsg += arr[i];
			strMsg += QString::fromLocal8Bit("��");
		}
	}
	if(!bRet)
	{
		strMsg = strMsg.mid(0,strMsg.size()-1);
		strMsg += QString::fromLocal8Bit("�ֶ�");
		QMessageBox::information(0, QString::fromLocal8Bit("��ʾ"), strMsg,QMessageBox::Ok);
	}

	return bRet;
}