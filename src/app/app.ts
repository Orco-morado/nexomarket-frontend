import { Component, OnInit, OnDestroy, inject, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterOutlet } from '@angular/router';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { InteractionStatus } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';
import { environment } from '../environments/environment';

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

  protected readonly title = 'NexoraMarket';
  readonly isLoggedIn = signal(false);
  readonly userName = signal<string | undefined>(undefined);

  ngOnInit(): void {
    this.msalService.handleRedirectObservable().subscribe({
      next: () => this.setLoginDisplay(),
      error: (error) => console.error('[MSAL] Error en el redirect:', error),
    });

    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this.destroying$)
      )
      .subscribe(() => this.setLoginDisplay());
  }

  setLoginDisplay(): void {
    const accounts = this.msalService.instance.getAllAccounts();
    if (accounts.length > 0 && !this.msalService.instance.getActiveAccount()) {
      this.msalService.instance.setActiveAccount(accounts[0]);
    }
    this.isLoggedIn.set(accounts.length > 0);
    this.userName.set(accounts[0]?.name);
  }

  login(): void {
    this.msalService.loginRedirect({ scopes: [environment.azure.apiScope] });
  }

  registrarse(): void {
    this.msalService.loginRedirect({
      scopes: [environment.azure.apiScope],
      prompt: 'create',
    });
  }

  logout(): void {
    this.msalService.logoutRedirect();
  }

  ngOnDestroy(): void {
    this.destroying$.next();
    this.destroying$.complete();
  }
}