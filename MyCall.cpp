#include <iostream>
#include <string>
#include <Windows.h>
#include <stdio.h>
#include <conio.h>
using namespace std;

double whoWins(double a,double b,double c,double d){
    if(a>b)
      if(a>c)
        if(a>d) return a;
        else return d;
      else if(c>d) return c;
         else return d;
    else if(b>c)
          if(b>d) return b;
          else return d;
         else if(c>d) return c;
              else return d;
}

double pointsCalc(int bids, int wins){
    double points;
    if(bids>wins) points=(double)(2*wins-bids);
    else if(bids==wins) points=(double)wins;
    else if(bids<wins) points=(double)(bids+((wins-bids)/(double)10));
    return points;
}

int main(){
    string players[4];
    char decision='Y';
    char cheat='Y';
    int cheater;
    int bids[4],wins[4];
    double points[4];
    cout<<"Welcome to Call Break Game";
    Sleep(1000);cout<<"!";Sleep(1000);cout<<"!\n";Sleep(1000);
    cout<<"\nWho are playing: ";
    for (int i=0; i<4; i++) cin>>players[i]; Sleep(1000);
    do {
        cout<<"\n\nPlace your bids in order("<<players[0]<<"/"<<players[1]<<"/"<<players[2]<<"/"<<players[3]<<"): ";
        for (int i=0; i<4; i++) cin>>bids[i];
        cout<<"\n\nOkay! Enjoy the game. (press any key to continue)";printf("%c", getch());
        cout<<"\n\n";
        do{ 
        cout<<"Place your wins in order("<<players[0]<<"/"<<players[1]<<"/"<<players[2]<<"/"<<players[3]<<"): ";
        for (int i=0; i<4; i++) cin>>wins[i];
        cout<<"\n\n";
        if(wins[0]+wins[1]+wins[2]+wins[3]!=13)
        cout<<"Something\'s wrong, "; } while (wins[0]+wins[1]+wins[2]+wins[3]!=13);
        for (int i=0; i<4; i++) points[i]=pointsCalc(bids[i],wins[i]);
        do {
        cout<<"did anyone cheat?(Y/N): "; cin>>cheat;
        if(cheat=='Y'||cheat=='y'){ cout<<"\n\nWho? ("<<players[0]<<":1,"<<players[1]<<":2,"<<players[2]<<":3,"<<players[3]<<":4): ";
        cin>>cheater;
        points[cheater-1]=0;
        cout<<"\n\nOther than that...";}} while(cheat=='Y'|| cheat=='y');
        cout<<"\n\nPlayers\t\t"<<"Bids\t\t"<<"Wins\t\t"<<"Points"<<endl;
        for (int i=0; i<4; i++) cout<<players[i]<<"\t\t"<<bids[i]<<"\t\t"<<wins[i]<<"\t\t"<<points[i]<<endl;
        int x; double winn=whoWins(points[0],points[1],points[2],points[3]);
        for (int i=0; i<4; i++) if(points[i]==winn) x=i;
        Sleep(1000);
        cout<<"\nCongratulations "<<players[x]<<"! You win this one with "<<winn<<" points"<<endl;
        Sleep(1000);
        cout<<"\nWanna play another round?(Y/N): "; cin>>decision;
    } while(decision!='N' && decision!='n');
    cout<<"\n\nRound ending in 5"; Sleep(1000);cout<<" 4";Sleep(1000);cout<<" 3";Sleep(1000);cout<<" 2";Sleep(1000);cout<<" 1";Sleep(1000);
    cout<<"\n\nThe round is over. Thanks for playing!";
    return 0;
}
