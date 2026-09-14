import { Component, OnInit, OnDestroy, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet } from '@angular/router';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { InteractionStatus } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';

@Component({
  imports: [CommonModule, RouterOutlet],
  selector: 'app-root',
  styleUrl: './app.css',
  templateUrl: './app.html',
})
export class App implements OnInit, OnDestroy {
  private readonly msalService = inject(MsalService);
  private readonly msalBroadcastService = inject(MsalBroadcastService);
  private readonly destroying$ = new Subject<void>();
  private heartbeat: ReturnType<typeof setInterval> | undefined;

  protected readonly title = 'NexoraMarket';
  isLoggedIn = false;
  userName: string | undefined;

  constructor() {
    console.log('%c[APP] Componente App CONSTRUIDO', 'color: orange; font-weight: bold;');
  }

  ngOnInit(): void {
    this.msalService.handleRedirectObservable().subscribe({
      next: (result) => {
        console.log('[MSAL] handleRedirectObservable resultado:', result);
        this.setLoginDisplay();
      },
      error: (error) => {
        console.error('[MSAL] Error en handleRedirectObservable:', error);
      }
    });

    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this.destroying$)
      )
      .subscribe(() => {
        const cuentas = this.msalService.instance.getAllAccounts();
        console.log('[MSAL] inProgress$ = None. Cuentas encontradas:', cuentas.length, cuentas);
        this.setLoginDisplay();
      });

    this.heartbeat = setInterval(() => {
      console.log('%c[HEARTBEAT] isLoggedIn =', 'color: dodgerblue;', this.isLoggedIn, '| userName =', this.userName);
    }, 2000);
  }

  setLoginDisplay(): void {
    const accounts = this.msalService.instance.getAllAccounts();
    this.isLoggedIn = accounts.length > 0;
    this.userName = accounts.length > 0 ? accounts[0].name : undefined;
    console.log('[MSAL] setLoginDisplay ejecutado. isLoggedIn =', this.isLoggedIn, '| userName =', this.userName);
  }

  login(): void {
    this.msalService.loginRedirect();
  }

  logout(): void {
    this.msalService.logoutRedirect();
  }

  ngOnDestroy(): void {
    console.log('%c[APP] Componente App DESTRUIDO', 'color: red; font-weight: bold;');
    if (this.heartbeat) {
      clearInterval(this.heartbeat);
    }
    this.destroying$.next(undefined);
    this.destroying$.complete();
  }
}