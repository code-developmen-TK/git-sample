Option Compare Database
Option Explicit
'------------------------------------------------------------------------------
'�V�X�e���g�p����萔�̒�`���W���[��
'-------------------------------------------------------------------------------

'ADO�@�J�[�\���^�C�v�萔
Public Const adOpenForwardOnly As Integer = 0 '�O���X�N���[���J�[�\���@���R�[�h�Z�b�g�̐擪���疖���Ɍ������Ĉړ����邱�Ƃ��ł���
Public Const adOpenKeyset As Integer = 1 '�L�[�Z�b�g�J�[�\���@���R�[�h�Z�b�g�̑S�Ă̕����Ɉړ����邱�Ƃ��ł���B���̃��[�U�[���X�V�������R�[�h�͎Q�Ƃ��邱�Ƃ��ł��܂����A�ǉ��A�폜�������R�[�h�͎Q�Ƃł��Ȃ��B
Public Const adOpenDynamic As Integer = 2 '���I�J�[�\���@���R�[�h�Z�b�g�̑S�Ă̕����Ɉړ����邱�Ƃ��ł���B���̃��[�U�[���ǉ��A�X�V�A�폜�������R�[�h���Q�Ƃ��邱�Ƃ��ł���B
Public Const adOpenStatic As Integer = 3 '�ÓI�J�[�\���@���R�[�h�Z�b�g�̑S�Ă̕����Ɉړ����邱�Ƃ��ł���B���̃��[�U�[�ɂ��ǉ��A�X�V�A�폜�͎Q�Ƃ��邱�Ƃ��ł��Ȃ��B

'ADO�@���b�N�^�C�v�萔
Public Const adLockReadOnly As Integer = 0 '�ǂݎ���p �f�[�^�̍X�V�E�ǉ��E�폜�͂ł��Ȃ�
Public Const adLockPessimistic As Integer = 1 '�r���I���b�N�@�ҏW����Ƀ��R�[�h�����b�N
Public Const adLockOptimistic As Integer = 2 '���L�I���b�N�@Update���\�b�h���Ăяo�����ꍇ�ɂ̂݁A���L�I���b�N
Public Const adLockBatchOptimistic As Integer = 3 '�����̃��R�[�h���o�b�`�X�V

'ADO �I�v�V�����萔
Public Const adCmdText As Integer = 1 '�R�}���h�܂��̓X�g�A�h �v���V�[�W���̃e�L�X�g��`�Ƃ��ĕ]�����܂��B

'ADO �J�[�\�����P�[�V�����萔
Public Const adUseClient As Integer = 3 '���[�J�� �J�[�\��
Public Const adUseServer As Integer = 2 '�f�[�^�v���o�C�_�[�J�[�\��

'ADO �X�L�[�}���萔
Public Const adSchemaColumns As Integer = 4      '�J�����̒�`
Public Const adSchemaTables As Long = 20         '�e�[�u���̒�`
Public Const adSchemaPrimaryKeys As Integer = 28 '��L�[�̒�`
 
'ADO �f�[�^�^�萔
Public Const adBoolean As Integer = 11         '�^�U�^
Public Const adUnsignedTinyInt As Integer = 17 '�o�C�g�^�i�����Ȃ��j
Public Const adSmallInt As Integer = 2         '�����^�i�����t���j
Public Const adInteger As Integer = 3          '�������^�i�����t���j
Public Const adCurrency As Integer = 6         '�ʉ݌^�i�����t���j
Public Const adSingle As Integer = 4           '�P���x���������_�^
Public Const adDouble As Integer = 5           '�{���x���������_�^
Public Const adDate As Integer = 7             '���t/�����^
Public Const adWChar As Integer = 130          '������^
Public Const adLongVarBinary As Integer = 205  '�����O�o�C�i���^

'ADO���R�[�h�Z�b�g�̏�Ԃ�\���萔
Public Const adStateClosed As Integer = 0     '�I�u�W�F�N�g�͕��Ă��邱�Ƃ������B
Public Const adStateOpen As Integer = 1       '�I�u�W�F�N�g�͊J���Ă��邱�Ƃ������B
Public Const adStateConnecting As Integer = 2 '�I�u�W�F�N�g�͐ڑ����Ă��邱�Ƃ������B
Public Const adStateExecuting As Integer = 4  '�I�u�W�F�N�g�̓R�}���h�����s���Ă��邱�Ƃ������B
Public Const adStateFetching As Integer = 8   '�I�u�W�F�N�g�̍s���擾����Ă��邱�Ƃ������B

'Office�I�u�W�F�N�g �t�@�C���E�t�H���_�I���_�C�A���O�Ŏg���萔
Public Const msoFileDialogFilePicker As Integer = 3 '�t�@�C����I������ꍇ
Public Const msoFileDialogFolderPicker As Integer = 4 '�t�H���_��I������ꍇ

'�e�X�g�f�[�^�V�X�e���S�̂Ŏg���ݒ�l��萔�ɃZ�b�g
Public Const SYSTEM_FILE_NAME As String = "�e�X�g�f�[�^�V�X�e��.accdb"
Public Const CONNECT_DATABASE_NAME As String = "�e�X�g�f�[�^�x�[�X.accdb"
Public Const DEFAULT_TEXT_FOLDER_PATH As String = "D:\VBA�J��\access\�e�L�X�g�f�[�^"
'Public Const DATABASE_PROVIDOR As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=D:\VBA�J��\access\�e�X�g�f�[�^�x�[�X.accdb"

'�V�X�e���Ŏg�p����e�[�u�����̒萔
'Public Const TEMPORARY_TABLE_NAME As String = "T_��������f�[�^�ǂݍ��ݗp�e�[�u��"
'Public Const LOADING_TABLE_NAME As String = "T_��������f�[�^�e�[�u��"
Public Const ACCEPTANCE_DATA_DABLE As String = "T_������f�[�^�e�[�u��"
Public Const ACCEPTANCE_INSPECT_DATA_DABLE As String = "T_����������f�[�^�e�[�u��"
'Public Const ACCEPTANCE_INSPECT_DATA_DABLE As String = "T_������O���e��Ή��f�[�^�e�[�u��"
Public Const ACCEPTANCE_INNER_OUTER_CORRESPONDENCE_TABLE As String = "T_������O���e��Ή��f�[�^�e�[�u��"
Public Const ACCEPTANCE_COMPOSITION_DATA_TABLE As String = "T_������g���f�[�^�e�[�u��"
Public Const ACCEPTANCE_SLIP_MANAGEMENT_TABLE As String = "T_������`�[�Ǘ�"
Public Const ACCEPTANCE_SLIP_DETALS_TABLE As String = "T_������`�[�ڍ׏��f�[�^�e�[�u��"
Public Const TREATMENT_TABLE As String = "T_�����e�[�u��"
Public Const TREATMENT_DATE_CORRESPOMDENCE As String = "T_�������Ή��e�[�u��"

'Excel����֌W�̒萔
Public Const xlDown As Integer = -4121                         '����
Public Const xlToLeft As Integer = 4159                        '����
Public Const xlToRight As Integer = 4161                       '�E��
Public Const xLUp As Integer = -4162                           '���

Public Const HISTORY_SHEET_FIRST_ROW As Long = 8               '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̍ŏ��̍s�B
Public Const HISTORY_SHEET_FIRST_CLUMN As Long = 1             '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̍ŏ��̗�B
Public Const HISTORY_SHEET_CLUMNS As Long = 34                 '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗񐔁B
Public Const HISTORY_SHEET_NUMBER1 As String = "�J�e�S���P"    '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�P�̖��O
Public Const HISTORY_SHEET_NUMBER2 As String = "�J�e�S��2"     '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�Q�̖��O
Public Const SPRIT_ROW As Long = 21                            '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�i�z��j�̍��E�����ʒu

'Public Const SOTOYOUKI_NUMBAER As Long = 4                     '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̊O�e��ԍ��̓����Ă����

'Public Const UCHIYOUKI_NUMBER1 As Long = 9                     '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̓��e��ԍ��̓����Ă����
'Public Const CONTENT1 As Long = 10                             '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̓��e���̓����Ă����
'Public Const WEIGHT1 As Long = 12                              '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̏d�ʂ̓����Ă����
'Public Const DOSE1 As Long = 13                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����
'Public Const ORENGE1 As Long = 14                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����
'Public Const GREEN1 As Long = 15                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����
'Public Const BLACK1 As Long = 16                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����

'Public Const UCHIYOUKI_NUMBER2 As Long = 21                    '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̓��e��ԍ��̓����Ă����
'Public Const CONTENT2 As Long = 24                             '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̓��e���̓����Ă����
'Public Const WEIGHT2 As Long = 23                              '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̏d�ʂ̓����Ă����
'Public Const DOSE2 As Long = 25                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����
'Public Const ORENGE2 As Long = 26                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����
'Public Const GREEN2 As Long = 27                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����
'Public Const BLACK2 As Long = 28                                '�ǂݍ��ރf�[�^�̓����Ă���V�[�g�̗ʂ̓����Ă����


Public Const DEFAULT_FOLDER As String = "D:\�v���O�������J��\excel\�����Ǘ��f�[�^" '�ŏ��ɊJ���t�H���_���w��
'Public Const PROCESSING_DATE As Long = 32 '�������̓�������
'Public Const TREATMENT_DATE_COLUMN As Long = 32 '�������̓�������
 
'Excel�����Ǘ��f�[�^�̗�ԍ��̒萔
Public Const cst�ʐ� As String = 1
Public Const cst�L�� As String = 2
Public Const cst�ԍ� As String = 3
Public Const cst�O�e��ԍ� As String = 4
Public Const cst������ As String = 5
Public Const cstW�� As String = 6
Public Const cst���[�� As String = 7
Public Const cst���� As String = 8
Public Const cst���e��ԍ�1 As String = 9
Public Const cst���e��1 As String = 10
Public Const cst��� As String = 11
Public Const cst�d��1 As String = 12
Public Const cst����1 As String = 13
Public Const cst�I�����W1 As String = 14
Public Const cst�~�h��1 As String = 15
Public Const cst�N��1 As String = 16
Public Const cst�O���� As String = 17
Public Const cst���� As String = 18
Public Const cst�߂� As String = 19
Public Const cst������ As String = 20
Public Const cst���e��ԍ�2 As String = 21
Public Const cst���� As String = 22
Public Const cst�d��2 As String = 23
Public Const cst���e��2 As String = 24
Public Const cst����2 As String = 25
Public Const cst�I�����W2 As String = 26
Public Const cst�~�h��2 As String = 27
Public Const cst�N��2 As String = 28
Public Const cst������ As String = 29
Public Const cst�u�����N As String = 30
Public Const cst�ۗ� As String = 31
Public Const cst������ As String = 32
Public Const cst�������o�b�`�ԍ� As String = 33
Public Const cst���l As String = 34
Public Const cst������̓��e��ԍ��ʒu = 1

'�g�p�e�[�u�����̈ꗗ
'MT_���
'MT_�ꏊ
'MT_���e����
'T_������`�[�Ǘ�
'T_������g��
'T_������`�[�ڍ׏��
'T_���������
'T_������e��Ή�
'T_��������
'T_����
'T_�������Ή�
'T_�����L�^
'T_���������[���
'T_�������O�e�핕��
'T_�������e����
'T_���o�e����