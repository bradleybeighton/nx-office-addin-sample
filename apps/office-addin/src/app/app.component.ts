import { Component, OnInit, OnDestroy } from '@angular/core';
import { MsalBroadcastService, MsalService } from '@azure/msal-angular';
import {
  AuthenticationResult,
  EventMessage,
  EventType
} from '@azure/msal-browser';
import { Router } from '@angular/router';
import { filter } from 'rxjs/operators';
import { Subscription } from 'rxjs';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit, OnDestroy {
  subs: Subscription[] = [];

  constructor(
    public router: Router,
    private msalService: MsalService,
    private msalBroadcastService: MsalBroadcastService
  ) {}

  ngOnInit() {
    this.subs.push(this.msalService.handleRedirectObservable().subscribe({
      next: (result: AuthenticationResult) => {
        if (result) {
        }
      },
      error: (error) => console.log(error)
    }));

   this.subs.push(this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS)
      )
      .subscribe((result: EventMessage) => {
        if (result && result.eventType === EventType.LOGIN_SUCCESS) {
         
        }
      }));
  }

  ngOnDestroy() {
    this.subs.map((sub: Subscription) => sub.unsubscribe());
  }
}