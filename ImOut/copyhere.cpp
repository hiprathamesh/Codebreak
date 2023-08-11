#include <libxl.h>
#include <iostream>
#include <string>
#include <cstring>
#include <Windows.h>
#include <stdio.h>
#include <conio.h>
#include <codecvt>
#include <locale>

double whoWins(double a, double b, double c, double d) {
    if (a > b)
        if (a > c)
            if (a > d) return a;
            else return d;
        else if (c > d) return c;
        else return d;
    else if (b > c)
        if (b > d) return b;
        else return d;
    else if (c > d) return c;
    else return d;
}

double pointsCalc(int bids, int wins) {
    double points;
    if (bids > wins) points = (double)(2 * wins - bids);
    else if (bids == wins) points = (double)wins;
    else if (bids < wins) points = (double)(bids + ((wins - bids) / (double)10));
    return points;
}

using namespace std;

int main() {
    wchar_t players[4][20]{}, date[20], note[100];
    char decision = 'y', cheat = 'y';
    int cheater, bids[4]{}, wins[4]{};
    double points[4]{}, total[4]{0};
    libxl::Book* book = xlCreateXMLBook();
    book->load(L"MyCall.xlsx");
    libxl::Sheet* sheet = book->getSheet(0);
    libxl::Sheet* sheet2 = book->getSheet(1);
    libxl::Font* font = book->addFont();
    font->setBold(true);
    libxl::Format* format = book->addFormat();
    format->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    format->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* boldFormat = book->addFormat();
    boldFormat->setFont(font);
    boldFormat->setBorderRight(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    boldFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    boldFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* tborderFormat = book->addFormat();
    tborderFormat->setBorderTop(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    tborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER); 
    tborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* bborderFormat = book->addFormat();
    bborderFormat->setBorderBottom(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    bborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    bborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* lborderFormat = book->addFormat();
    lborderFormat->setBorderLeft(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    lborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    lborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* rborderFormat = book->addFormat();
    rborderFormat->setBorderRight(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    rborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    rborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* tlborderFormat = book->addFormat();
    tlborderFormat->setBorderTop(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    tlborderFormat->setBorderLeft(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    tlborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    tlborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* trborderFormat = book->addFormat();
    trborderFormat->setBorderTop(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    trborderFormat->setBorderRight(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    trborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    trborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* blborderFormat = book->addFormat();
    blborderFormat->setBorderBottom(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    blborderFormat->setBorderLeft(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    blborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    blborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    libxl::Format* brborderFormat = book->addFormat();
    brborderFormat->setBorderBottom(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    brborderFormat->setBorderRight(libxl::BorderStyle::BORDERSTYLE_MEDIUM);
    brborderFormat->setAlignH(libxl::AlignH::ALIGNH_CENTER);
    brborderFormat->setAlignV(libxl::AlignV::ALIGNV_CENTER);
    int lastRow = sheet->lastRow();
    int lastRow2 = sheet2->lastRow();
    cout << "welcome to call break game";
    Sleep(500); cout << "!"; Sleep(500); cout << "!\n"; Sleep(500);
    cout << "\n\nWhat\'s the date today? ";
    wcin.getline(date,11); Sleep(1000);
    sheet->writeStr(lastRow , 4, L"", tlborderFormat);
    sheet->writeStr(lastRow + 1, 4, date, blborderFormat );
    sheet->writeStr(lastRow, 5, L"", tlborderFormat);
    sheet->writeStr(lastRow + 1, 5, L"", blborderFormat);
    sheet->writeStr(lastRow, 6, L"", tlborderFormat);
    sheet->writeStr(lastRow, 7, L"", tborderFormat);
    sheet->writeStr(lastRow, 10, L"", tborderFormat);
    sheet->writeStr(lastRow, 13, L"", tborderFormat);
    sheet->writeStr(lastRow, 16, L"", tborderFormat);
    sheet->writeStr(lastRow, 19, L"", tborderFormat);
    sheet->writeStr(lastRow, 8, L"", tlborderFormat);
    sheet->writeStr(lastRow, 9, L"", tlborderFormat);
    sheet->writeStr(lastRow, 12, L"", tlborderFormat);
    sheet->writeStr(lastRow, 11, L"", tlborderFormat);
    sheet->writeStr(lastRow, 14, L"", tlborderFormat);
    sheet->writeStr(lastRow, 15, L"", tlborderFormat);
    sheet->writeStr(lastRow, 17, L"", tlborderFormat);
    sheet->writeStr(lastRow, 18, L"", tlborderFormat);
    sheet->writeStr(lastRow, 20, L"", trborderFormat);
    sheet->writeStr(lastRow + 1, 8, L"", blborderFormat);
    sheet->writeStr(lastRow + 1, 11, L"", blborderFormat);
    sheet->writeStr(lastRow + 1, 14, L"", blborderFormat);
    sheet->writeStr(lastRow + 1, 17, L"", blborderFormat);
    sheet->writeStr(lastRow + 1, 18, L"", blborderFormat);
    sheet->writeStr(lastRow + 1, 20, L"", brborderFormat);
    sheet->writeStr(lastRow + 1, 19, L"", bborderFormat);
    cout << "\n\nWho are playing:\n";
    wcin.getline(players[0], 20); wcin.getline(players[1], 20);wcin.getline(players[2], 20);wcin.getline(players[3], 20);
    sheet->setMerge(lastRow + 1, lastRow + 1, 6, 7);
    sheet->setMerge(lastRow + 1, lastRow + 1, 9, 10);
    sheet->setMerge(lastRow + 1, lastRow + 1, 12, 13);
    sheet->setMerge(lastRow + 1, lastRow + 1, 15, 16);
    for (int i = 0; i < 4;i++) sheet->writeStr(lastRow + 1, (i + 2) * 3, players[i], blborderFormat);
    sheet->writeStr(lastRow + 1, 7, L"", bborderFormat);
    sheet->writeStr(lastRow + 1, 10, L"", bborderFormat);
    sheet->writeStr(lastRow + 1, 13, L"", bborderFormat);
    sheet->writeStr(lastRow + 1, 16, L"", bborderFormat);
    for (int i = 0; i < ((lastRow2 - 11) / 2); i++) {
        for (int j = 0; j < 4; j++) {
            if (wcscmp(sheet2->readStr(2 * i + 11, 4), players[j]) == 0) {
                double oldMatch = sheet2->readNum(2 * i + 11, 6);
                double newMatch = oldMatch + 1;
                sheet2->writeNum(2 * i + 11, 6, newMatch);
            }
        }
    }
    double srno = 1;
    do {
        int lastRow = sheet->lastRow();
        sheet->writeNum(lastRow + 1, 5, srno, lborderFormat);
        wcout << "\n\nplace your bids in order(" << players[0] << "/" << players[1] << "/" << players[2] << "/" << players[3] << "): ";
        for (int i = 0; i < 4; i++) cin >> bids[i];
        sheet->writeNum(lastRow + 1, 6, bids[0], lborderFormat);
        sheet->writeNum(lastRow + 1, 9, bids[1], lborderFormat);
        sheet->writeNum(lastRow + 1, 12, bids[2], lborderFormat);
        sheet->writeNum(lastRow + 1, 15, bids[3], lborderFormat);

        sheet->writeStr(lastRow, 4, L"", lborderFormat);
        sheet->writeStr(lastRow + 1, 4, L"", lborderFormat);
        sheet->writeStr(lastRow + 2, 4, L"", lborderFormat);
        sheet->writeStr(lastRow, 5, L"", lborderFormat);
        sheet->writeStr(lastRow + 2, 5, L"", lborderFormat);
        sheet->writeStr(lastRow + 2, 17, L"", rborderFormat);
        sheet->writeStr(lastRow + 1, 21, L"", lborderFormat);
        sheet->writeStr(lastRow + 2, 21, L"", lborderFormat);
        sheet->writeStr(lastRow, 6, L"", lborderFormat);
        sheet->writeStr(lastRow, 7, L"", rborderFormat);
        sheet->writeStr(lastRow, 9, L"", lborderFormat);
        sheet->writeStr(lastRow, 10, L"", rborderFormat);
        sheet->writeStr(lastRow, 12, L"", lborderFormat);
        sheet->writeStr(lastRow, 13, L"", rborderFormat);
        sheet->writeStr(lastRow, 15, L"", lborderFormat);
        sheet->writeStr(lastRow, 16, L"", rborderFormat);
        sheet->writeStr(lastRow, 18, L"", lborderFormat);
        sheet->writeStr(lastRow, 20, L"", rborderFormat);

        cout << "\n\nokay! enjoy the game. (press any key to continue)"; printf("%c", _getch());
        cout << "\n\n\n";
        do {
            wcout << "place your wins in order(" << players[0] << "/" << players[1] << "/" << players[2] << "/" << players[3] << "): ";
            for (int i = 0; i < 4; i++) cin >> wins[i];
            cout << "\n\n";
            if (wins[0] + wins[1] + wins[2] + wins[3] != 13)
                cout << "something\'s wrong, ";
        } while (wins[0] + wins[1] + wins[2] + wins[3] != 13);
        sheet->writeNum(lastRow + 1, 7, wins[0], rborderFormat);
        sheet->writeNum(lastRow + 1, 10, wins[1], rborderFormat);
        sheet->writeNum(lastRow + 1, 13, wins[2], rborderFormat);
        sheet->writeNum(lastRow + 1, 16, wins[3], rborderFormat);
        for (int i = 0; i < 4; i++) points[i] = pointsCalc(bids[i], wins[i]);
        for (int i = 0; i < 4; i++) {
            if (wins[i] >= 2 * bids[i]) points[i] = 0;
        }
        do {
            cout << "did anyone cheat?(y/n): "; cin >> cheat;
            if (cheat == 'Y' || cheat == 'y') {
                wcout << "\n\nwho? (" << players[0] << ":1," << players[1] << ":2," << players[2] << ":3," << players[3] << ":4): ";
                cin >> cheater;
                points[cheater - 1] = 0;
                cout << "\n\nother than that...";
            }
        } while (cheat == 'Y' || cheat == 'y');
        sheet->writeNum(lastRow + 2, 6, points[0], lborderFormat);
        sheet->writeNum(lastRow + 2, 9, points[1], lborderFormat);
        sheet->writeNum(lastRow + 2, 12, points[2], lborderFormat);
        sheet->writeNum(lastRow + 2, 15, points[3], lborderFormat);
        for (int i = 0; i < 4; i++) total[i] += points[i];
        sheet->writeNum(lastRow + 2, 7, total[0], rborderFormat);
        sheet->writeNum(lastRow + 2, 10, total[1], rborderFormat);
        sheet->writeNum(lastRow + 2, 13, total[2], rborderFormat);
        sheet->writeNum(lastRow + 2, 16, total[3], rborderFormat);
        cout << "\n\nplayers\t\t\t" << "bids\t\t\t" << "wins\t\t\t" << "points" << endl;
        for (int i = 0; i < 4; i++) wcout << players[i] << "\t\t\t" << bids[i] << "\t\t\t" << wins[i] << "\t\t\t" << points[i] << endl;
        int x; double winn = whoWins(points[0], points[1], points[2], points[3]);
        for (int i = 0; i < 4; i++) if (points[i] == winn) x = i;
        Sleep(1000);
        wcout << "\ncongratulations " << players[x] << "! you win this one with " << winn << " points" << endl;
        Sleep(1000);
        cout << "\n\nany interesting note about this match: ";
        wcin.ignore();
        wcin.getline(note, 100);
        sheet->setMerge(lastRow + 1, lastRow + 2, 18, 20);
        sheet->writeStr(lastRow + 1, 18, note, lborderFormat);
        for (int i = 0; i < ((lastRow2 - 11) / 2); i++) {
            for (int j = 0; j < 4; j++) {
                if (wcscmp(sheet2->readStr(2*i+11,4),players[j])==0) {
                    double oldRound = sheet2->readNum(2 * i + 11, 7);
                    double newRound = oldRound + 1;
                    sheet2->writeNum(2 * i + 11, 7, newRound);

                    double oldBid = sheet2->readNum(2 * i + 11, 8);
                    double newBid = oldBid + bids[j];
                    sheet2->writeNum(2 * i + 11, 8, newBid);

                    double oldWin = sheet2->readNum(2 * i + 11, 9);
                    double newWin = oldWin + wins[j];
                    sheet2->writeNum(2 * i + 11, 9, newWin);

                    double oldPoint = sheet2->readNum(2 * i + 11, 10);
                    double newPoint = oldPoint + points[j];
                    sheet2->writeNum(2 * i + 11, 10, newPoint);

                    double oldHigh = sheet2->readNum(2 * i + 11, 12);
                    if (points[j] >= oldHigh)
                        sheet2->writeNum(2 * i + 11, 12, points[j]);

                    double oldLow = sheet2->readNum(2 * i + 11, 13);
                    if (points[j] <= oldLow)
                        sheet2->writeNum(2 * i + 11, 13, points[j]);
                }
            }
        }
        cout << "\n\nwanna play another round?(y/n): "; cin >> decision;
        srno += 1;
    } while (decision != 'N' && decision != 'n');
    int laastRow = sheet->lastRow();
    sheet->writeStr(laastRow, 4, L"", lborderFormat);
    sheet->writeStr(laastRow, 5, L"", lborderFormat);
    sheet->writeStr(laastRow, 6, L"", lborderFormat);
    sheet->writeStr(laastRow, 7, L"", rborderFormat);
    sheet->writeStr(laastRow, 9, L"", lborderFormat);
    sheet->writeStr(laastRow, 10, L"", rborderFormat);
    sheet->writeStr(laastRow, 12, L"", lborderFormat);
    sheet->writeStr(laastRow, 13, L"", rborderFormat);
    sheet->writeStr(laastRow, 15, L"", lborderFormat);
    sheet->writeStr(laastRow, 16, L"", rborderFormat);
    sheet->writeStr(laastRow, 18, L"", lborderFormat);
    sheet->writeStr(laastRow, 20, L"", rborderFormat);
    sheet->writeStr(laastRow + 1, 4, L"", lborderFormat);
    sheet->writeStr(laastRow + 1, 5, L"", lborderFormat);
    sheet->writeStr(laastRow + 1, 18, L"", lborderFormat);
    sheet->writeStr(laastRow + 1, 20, L"", rborderFormat);
    double largesttot[4]{ 0 };
    for (int i = 0; i < 4; i++) {
        if (total[i] > largesttot[0]) largesttot[0] = total[i];
    }
    for (int i = 0; i < 4; i++) {
        if (total[i] != largesttot[0]) {
            if (total[i] > largesttot[1]) largesttot[1] = total[i];
        }
    }
    for (int i = 0; i < 4; i++) {
        if (total[i] != largesttot[0] && total[i] != largesttot[1]) {
            if (total[i] > largesttot[2]) largesttot[2] = total[i];
        }
    }
    for (int i = 0; i < 4; i++) {
        if (total[i] != largesttot[0] && total[i] != largesttot[1] && total[i] != largesttot[2]) largesttot[3] = total[i];
    }
    for (int i = 0; i < 4; i++) {
        for (int j = 0; j < 4; j++) {
                sheet->writeNum(laastRow + 1, (i + 2) * 3, total[i], lborderFormat);
            if (total[i] == largesttot[j]) {
                sheet->writeNum(laastRow + 1, ((i + 2) * 3) + 1, (j + 1), boldFormat);
            }
        }
    }
    sheet->writeStr(laastRow + 2, 4, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 5, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 6, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 7, L"", brborderFormat);
    sheet->writeStr(laastRow + 2, 9, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 10, L"", brborderFormat);
    sheet->writeStr(laastRow + 2, 12, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 13, L"", brborderFormat);
    sheet->writeStr(laastRow + 2, 15, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 16, L"", brborderFormat);
    sheet->writeStr(laastRow + 2, 18, L"", blborderFormat);
    sheet->writeStr(laastRow + 2, 20, L"", brborderFormat);
    sheet->writeStr(laastRow + 2, 8, L"", bborderFormat);
    sheet->writeStr(laastRow + 2, 11, L"", bborderFormat);
    sheet->writeStr(laastRow + 2, 14, L"", bborderFormat);
    sheet->writeStr(laastRow + 2, 17, L"", bborderFormat);
    sheet->writeStr(laastRow + 2, 19, L"", bborderFormat);
    cout << "\n\nplayers\t\t\t" << "total points" << endl;
    for (int i = 0; i < 4; i++) wcout << players[i] << "\t\t\t" << total[i] << endl;
    int y; double winnn = whoWins(total[0], total[1], total[2], total[3]);
    for (int i = 0; i < 4; i++) if (total[i] == winnn) y = i;
    Sleep(1000);
    wcout << "\ncongratulations " << players[y] << "! you win this match with " << winnn << " total points\n(press any key to continue)" << endl; printf("%c", _getch());
    Sleep(1000);
    cout << "\n\nmatch ending in 5"; Sleep(1000); cout << " 4"; Sleep(1000); cout << " 3"; Sleep(1000); cout << " 2"; Sleep(1000); cout << " 1"; Sleep(1000);
    cout << "\n\nthe match is over. thanks for playing!\n\n";
    book->save(L"MyCall.xlsx");
    book->release();
    return 0;
}
