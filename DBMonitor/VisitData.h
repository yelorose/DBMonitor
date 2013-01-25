// VisitData.h: interface for the CVisitData class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VISITDATA_H__901ECAF9_1220_4C2F_A47C_378412E500E1__INCLUDED_)
#define AFX_VISITDATA_H__901ECAF9_1220_4C2F_A47C_378412E500E1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define		DB_TYPE_SQLSERVER	1
#define		DB_TYPE_ACCESS		2

class CVisitData  
{
public:
	CVisitData();
	virtual ~CVisitData();

private:
	_RecordsetPtr m_pDBRecordSet;	// ָ�����ݿ�ģ�顱�ĵ�ǰ��¼����ָ��

	CString m_strDBServer;		//���ݷ���������
	CString m_strDBUserName;	//���ݿ������û�����
	CString m_strDBPassword;	//���ݿ���������
	CString m_strDBName;		//���ݿ�����
	int m_nDBType;				//���ݿ�����

protected:
	_ConnectionPtr m_pDBConn;		// ָ�����ݿ����Ӷ����ָ��
	_StreamPtr m_pStm;				// ָ�������������ָ��
	int m_nRecordCount;				// ��ǰ��¼���ļ�¼��

public:
	//��ʼ�����ݿ�ģ��
	BOOL Initialize(LPCTSTR lpszDBServer, LPCTSTR lpszDBUserName, LPCTSTR lpszDBPassword, LPCTSTR lpszDBName, int nDBType = DB_TYPE_SQLSERVER);
	//����ʼ�����ݿ�ģ��
	BOOL Uninitialize();
	
	BOOL BeginTrans();						//��ʼ�������
	BOOL CommitTrans();						//�ύ����
	BOOL RollbackTrans();					//����ִ��ʧ�ܣ��ع�

	BOOL Execute(LPCTSTR lpszSql);			//ִ��SQL����
	int  InsertForID(LPCTSTR lpszSql);		//�������ݣ���ȡ���Զ����ɵ�ID

	BOOL SearchData(LPCTSTR lpszSqlSel);	//��ɶ����ݵĲ�ѯ����
	long GetRecordCount();					//��ȡ��ѯ����ļ�¼��Ŀ
	BOOL RecordSetMoveNext();				//����ǰ��¼ָ������ƶ�һ����¼
	BOOL RecordSetMovePrev();				//����ǰ��¼ָ����ǰ�ƶ�һ����¼
	BOOL RecordSetMove(long lStepSize);		//�ƶ���ǰ��¼ָ��
	BOOL RecordSetMoveLast();				//�����һ����¼�趨Ϊ��¼���ĵ�ǰ��¼
	BOOL RecordSetMoveFirst();				//����һ����¼�趨Ϊ��¼���ĵ�ǰ��¼
	BOOL IsRecordSetEOF();					//��ȡ��ǰ��¼�Ƿ��Ǽ�¼�������һ����¼
	BOOL IsRecordSetBOF();					//��ȡ��ǰ��¼�Ƿ��Ǽ�¼���ĵ�һ����¼
	BOOL IsRecordSetEmpty();				//��ȡ��¼���Ƿ�Ϊ��
	_variant_t GetFieldValueOfRecordSet(LPCTSTR lpszFieldName);	//��ȡ��ǰ��¼����ĳ���ֶε�ֵ
	long GetFieldValue_Long(LPCTSTR lpszFieldName);				//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��long��
	short GetFieldValue_Short(LPCTSTR lpszFieldName);			//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��short��
	BYTE GetFieldValue_Byte(LPCTSTR lpszFieldName);				//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��BYTE��
	BOOL GetFieldValue_Bool(LPCTSTR lpszFieldName);				//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��BOOL��
	float GetFieldValue_Float(LPCTSTR lpszFieldName);			//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��float��
	double GetFieldValue_Double(LPCTSTR lpszFieldName);			//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��double��
	CString GetFieldValue_String(LPCTSTR lpszFieldName);		//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��string��
	COleDateTime GetFieldValue_DateTime(LPCTSTR lpszFieldName);	//��ȡ��ǰ��¼����ĳ���ֶε�ֵ��COleDateTime��
	BOOL SetFieldValueOfRecordSet(LPCTSTR lpszFieldName, _variant_t varValue);	//���õ�ǰ��¼����ĳ���ֶε�ֵ
	BOOL SetFieldValue_Long(LPCTSTR lpszFieldName, long lValue);				//���õ�ǰ��¼����ĳ���ֶε�ֵ��long��
	BOOL SetFieldValue_Short(LPCTSTR lpszFieldName, short nValue);				//���õ�ǰ��¼����ĳ���ֶε�ֵ��short��
	BOOL SetFieldValue_Byte(LPCTSTR lpszFieldName, BYTE byValue);				//���õ�ǰ��¼����ĳ���ֶε�ֵ��BYTE��
	BOOL SetFieldValue_Bool(LPCTSTR lpszFieldName, BOOL bValue);				//���õ�ǰ��¼����ĳ���ֶε�ֵ��BOOL��
	BOOL SetFieldValue_Float(LPCTSTR lpszFieldName, float fValue);				//���õ�ǰ��¼����ĳ���ֶε�ֵ��float��
	BOOL SetFieldValue_Double(LPCTSTR lpszFieldName, double dValue);			//���õ�ǰ��¼����ĳ���ֶε�ֵ��double��
	BOOL SetFieldValue_String(LPCTSTR lpszFieldName, LPCTSTR lpszValue);		//���õ�ǰ��¼����ĳ���ֶε�ֵ��string��
	BOOL SetFieldValue_DateTime(LPCTSTR lpszFieldName, COleDateTime otDateTime);//���õ�ǰ��¼����ĳ���ֶε�ֵ��COleDateTime��
	BOOL UpdateRecordSet();					//���µ�ǰ��¼��

protected:
};

#endif // !defined(AFX_VISITDATA_H__901ECAF9_1220_4C2F_A47C_378412E500E1__INCLUDED_)
