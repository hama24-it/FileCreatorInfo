#-----------------------------------------------------------
# ���ړI
# �E�t�@�C���̍쐬�ҏ����m�F���A�Ǘ��O�ō쐬���ꂽ�t�@�C�����Ȃ����m�F

# ���Ώۂ̏ꏊ
# �E�w�肵���t�H���_�i�T�u�t�H���_���܂ށj�̔z���ɂ���t�@�C��

# ���Ώۂ̃t�@�C���g���q
# �E"*.xls*"
# �E"*.doc*"

# ������݌v
# 01.�Ώۂ̃t�H���_�̏����擾
# 02.foreach ���[�v�ɂ��擾�����t�@�C��������������
# 03.�t�@�C���̏ꍇ�̏��������s
#	�Eif ��������ɂ��A�w�肵���t�@�C���g���q�̏ꍇ�ɏ��������s
#	�E�t�@�C���쐬�ҏ����擾���ϐ��Ɋi�[
# 04.�ϐ��Ɋi�[�������𐮌`�������ɕ\��
#-----------------------------------------------------------
# [Code]
# ���V�F���I�u�W�F�N�g���쐬 
$shell = New-Object -ComObject Shell.Application

# �������Ώۂ̃t�H���_
$targetFolder = "C:\Users\hama24\Desktop\test"

# ���t�@�C���g���q�̎w��
[string[]]$include = @("*.xls*", "*.doc*")

# ���ϐ��Ƀt�H���_�̏����i�[�i-Recurse�F�T�u�t�H���_���܂ށj
$itemList = Get-ChildItem $targetFolder -Recurse

foreach($item in $itemList)
{
    # ��PSIsContainer �v���p�e�B�ɂ��t�H���_�E�t�@�C���̔���
    if($item.PSIsContainer)
    {
        # ���t�H���_�̏ꍇ�̏���
	# 
    }
    else
    {
	# ���t�H���_�̎w��
	$shellFolder = $shell.NameSpace($item.Directory.FullName)

	# ���t�@�C���̎�ނ��w��
	$File_01 = (Get-ChildItem $shellFolder.parseName($item.Name).Path -include $include).Name

	# ��������NULL�E�󔒂�����
	if([string]::IsNullOrEmpty($File_01))
	{
	# ��̏ꍇ�̏���
	# 
	}
	else
	{
	# ���w�肵���t�@�C�������݂���ꍇ�̏���
	$shellFile = $shellFolder.parseName($File_01)

        # ���t�@�C���̏ꍇ�̏���
	# ���������ϐ��Ɋi�[
	$Create = "20(�쐬��)_"
	# ���t�@�C���̍쐬�ҏ���ϐ��Ɋi�[
	$CreateData = $shellFolder.GetDetailsOf($shellFile, 20)
	Write-output "$Create$CreateData,$item"
	}
    }
} 
#-----------------------------------------------------------
