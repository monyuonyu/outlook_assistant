@echo off

title Outlook�鏑�A�V�X�^���g
echo Outlook�鏑�A�V�X�^���g���N�����Ă��܂�...

:: Python�̉��z�����A�N�e�B�x�[�g
call .venv\Scripts\activate.bat

:: API KEY�����ϐ��Ƃ��Đݒ�
set ANTHROPIC_API_KEY=YOYR_API_KEY

:: �J�X�^���ݒ�ŃX�N���v�g�����s
python outlook_assistant.py ^
  --emails 15 ^
  --days 14 ^
  --priority-keywords ���},�d�v,�ً} ^
  --working-hours 0830 1730 ^
  --focus-time 10 12 ^
  --report-style detailed ^
  --api-key %ANTHROPIC_API_KEY%

:: ���s��ɃE�B���h�E���J�����܂܂ɂ���
pause
