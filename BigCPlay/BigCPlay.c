
#include <stdio.h>
#include <windows.h>
#include <mmsystem.h>
#pragma comment(lib, "winmm.lib")
#include <tchar.h>

SHORT keyCode[42];
HWND gameWindowHandle;
int midiDeviceIndex;
unsigned char startNote;

#define PM_NONE 0
#define PM_PLAYING 1
#define PM_SET_START_NOTE 2
int playMode;

SHORT translateKeyCode(char note) {
    note -= startNote;
    if (note < 37)
        return keyCode[note];
    else
        return 0;
}

void CALLBACK processMidiMessage(HMIDIIN handle, UINT messageType, DWORD dwInstance, DWORD dwParam1, DWORD dwParam2) {
    unsigned char type, note;
    switch (messageType) {
    case MIM_DATA:
        type = ((char*)&dwParam1)[0];
        note = ((char*)&dwParam1)[1];
        if (type == 144) {
            switch (playMode) {
            case PM_PLAYING:
                PostMessage(gameWindowHandle, WM_KEYDOWN, translateKeyCode(note), 0);
                Sleep(50);
                break;
            case PM_SET_START_NOTE:
                startNote = note;
                break;
            }
         }
        else if (type == 128) {
            if (playMode == PM_PLAYING)
                PostMessage(gameWindowHandle, WM_KEYUP, translateKeyCode(note), 0);
        }
        break;
    }
}

void pressQToExit() {
    printf("按Q键退出。\n");
    while (TRUE) {
        char c = getch();
        if ((c == 'q') || (c == 'Q'))
            break;
    }
}

void userSetKeySequence(char* keyMap, char* hint, int count, int* index) {
    char tmp[64];
    while (TRUE) {
        printf("%s（共%d个）：", hint, count);
        scanf_s("%s", tmp, sizeof(tmp));
        if (strlen(tmp) == count) {
            int i;
            for (i = 0; i < count; i++)
                keyMap[index[i]] = tolower(tmp[i]);
            break;
        }
        printf("个数不对，请重新输入。\n");
    }
}

BOOL loadKeyMap() {
    char keyMap[38];
    int keyMapIndexW1[] = { 24, 26, 28, 29, 31, 33, 35, 36 };
    int keyMapIndexW2[] = { 12, 14, 16, 17, 19, 21, 23 };
    int keyMapIndexW3[] = { 0, 2, 4, 5, 7, 9, 11 };
    int keyMapIndexB1[] = { 25, 27, 30, 32, 34 };
    int keyMapIndexB2[] = { 13, 15, 18, 20, 22 };
    int keyMapIndexB3[] = { 1, 3, 6, 8, 10 };
    FILE* fp = NULL;
    fopen_s(&fp, "keyMap.txt", "r");
    if (fp == NULL) {
        printf("需要设置键位，请打开FF14中全键盘键位设置窗口，对照输入，中间不要空格。\n");
        memset(keyMap, 32, 37);
        keyMap[37] = 0;
        userSetKeySequence(keyMap, "依次输入白键第一行 1 2 3 4 5 6 7 i 对应的键位", 8, keyMapIndexW1);
        userSetKeySequence(keyMap, "依次输入白键第二行 1 2 3 4 5 6 7 对应的键位", 7, keyMapIndexW2);
        userSetKeySequence(keyMap, "依次输入白键第三行 1 2 3 4 5 6 7 对应的键位", 7, keyMapIndexW3);
        userSetKeySequence(keyMap, "依次输入黑键第一行 1# 3b 4# 5# 7b 对应的键位", 5, keyMapIndexB1);
        userSetKeySequence(keyMap, "依次输入黑键第二行 1# 3b 4# 5# 7b 对应的键位", 5, keyMapIndexB2);
        userSetKeySequence(keyMap, "依次输入黑键第三行 1# 3b 4# 5# 7b 对应的键位", 5, keyMapIndexB3);
        playMode = PM_SET_START_NOTE;
        startNote = 0;
        printf("请按一下MIDI键盘上最左边的C键。\n");
        while (TRUE) {
            if (startNote > 0)
                break;
            Sleep(50);
        }
        playMode = PM_NONE;
        printf("已设置起始音符为%u。\n", (unsigned int)startNote);
        printf("键位设置完毕，如需重新设置，请删除keyMap.txt文件，再重新启动本软件。\n");
        fopen_s(&fp, "keyMap.txt", "w");
        if (fp == NULL) {
            printf("无法写入文件。\n");
            return FALSE;
        }
        fwrite(keyMap, 1, 37, fp);
        fprintf(fp, "\r\n%u", (unsigned int)startNote);
        fclose(fp);
    }
    else {
        int n;
        n = fread(keyMap, 1, 37, fp);
        if (n != 37) {
            printf("键位设置文件keyMap.txt不正确。\n");
            fclose(fp);
            return FALSE;
        }
        keyMap[37] = 0;
        unsigned int x;
        n = fscanf_s(fp, "%u", &x);
        if (n != 1) {
            printf("键位设置文件keyMap.txt不正确。\n");
            fclose(fp);
            return FALSE;
        }
        startNote = (unsigned char)x;
        fclose(fp);
    }
    printf("已读取键位设置：%s，起始音符：%u。\n", keyMap, (unsigned int)startNote);
    unsigned int i;
    for (i = 0; i < 37; i++)
        keyCode[i] = VkKeyScan(keyMap[i]);
    return TRUE;
}

BOOL chooseGameWindow() {
    HWND handle = NULL, windowHandles[8];
    int numWindows = 0;
    while ((handle = FindWindowEx(NULL, handle, NULL, _T("最终幻想XIV"))) != NULL) {
        windowHandles[numWindows] = handle;
        numWindows++;
        if (numWindows == 8)
            break;
    }
    if (numWindows == 0) {
        gameWindowHandle = 0;
        printf("没有找到FF14游戏窗口。\n");
    }
    else if (numWindows == 1) {
        gameWindowHandle = windowHandles[0];
        printf("选中了FF14游戏窗口%lu。\n", (unsigned long)gameWindowHandle);
        return TRUE;
    }
    else {
        printf("找到多个FF14游戏窗口：\n");
        int i;
        for (i = 0; i < numWindows; i++)
            printf("  %d: %lu\n", i + 1, (unsigned long)windowHandles[i]);
        while (TRUE) {
            printf("输入要选择的窗口序号：");
            scanf_s("%d", &i);
            i--;
            if ((i < 0) || (i >= numWindows)) {
                printf("序号不正确。\n");
                continue;
            }
            gameWindowHandle = windowHandles[i];
            printf("选中了FF14游戏窗口%lu。\n", (unsigned long)gameWindowHandle);
            return TRUE;
        }
    }
    return FALSE;
}

BOOL chooseMidiDevice() {
    int numDevices, i;
    numDevices = midiInGetNumDevs();
    if (numDevices == 0) {
        printf("没有MIDI输入设备。\n");
        return FALSE;
    }
    else if (numDevices == 1) {
        midiDeviceIndex = 0;
        printf("选中了唯一的MIDI输入设备。\n");
        return TRUE;
    }
    else {
        printf("找到%d个MIDI输入设备。\n", numDevices);
        while (TRUE) {
            printf("输入要选择的MIDI输入设备序号：");
            scanf_s("%d", &i);
            i--;
            if ((i < 0) || (i >= numDevices)) {
                printf("序号不正确。\n");
                continue;
            }
            midiDeviceIndex = i;
            printf("选中了%d号MIDI输入设备。\n", i + 1);
            return TRUE;
        }
    }
}

int main()
{
    printf("=== 克拉维亚·吹小曲儿 ===\n{拉诺西亚,龙巢神殿,延夏,银泪湖}的Clavy好胖！\n\n");

    if (!chooseMidiDevice()) {
        pressQToExit();
        return 1;
    }

    playMode = PM_NONE;
    HMIDIIN midiHandle;
    MMRESULT result = midiInOpen(&midiHandle, 0, (DWORD_PTR)processMidiMessage, 0, CALLBACK_FUNCTION);
    if (result != MMSYSERR_NOERROR) {
        printf("无法连接MIDI输入设备，错误代码：%u。\n", result);
        pressQToExit();
        return 4;
    }
    result = midiInStart(midiHandle);
    if (result != MMSYSERR_NOERROR) {
        midiInClose(midiHandle);
        printf("无法从MIDI输入设备读取数据，错误代码：%u。\n", result);
        pressQToExit();
        return 6;
    }

    if (!loadKeyMap()) {
        pressQToExit();
        return 2;
    }
    if (!chooseGameWindow()) {
        pressQToExit();
        return 3;
    }

    playMode = PM_PLAYING;
    printf("现在可以开始弹琴了。\n");
    pressQToExit();

    midiInStop(midiHandle);
    midiInReset(midiHandle);
    midiInClose(midiHandle);

    return 0;
}
