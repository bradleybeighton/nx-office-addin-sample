import { Component, OnInit } from '@angular/core';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import {
  RedirectRequest
} from '@azure/msal-browser';

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss']
})
export class LoginComponent implements OnInit {
  msalAuthProcess = false;

  constructor(
    public msalService: MsalService,
    public msalBroadcastService: MsalBroadcastService
  ) {}

  ngOnInit() {
    this.performSSO();
  }

  performSSO() {
    const options = {
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      };
    
   
    try {
      OfficeRuntime.auth
        .getAccessToken(options)
        .then(async (result: any) => {
          if (result) {
            // begin the behalf-of-flow to get MS Graph token
          }
        })
        .catch((err) => {
          this.msalAuthProcess = true;
        });
    } catch (err) {
      this.msalAuthProcess = true;
    }
  }

  performMSALRedirect() {
    const options = {
      forceRefresh: true,
      loginHint: '',
      prompt: 'select_account',
      scopes: ["your_defined_scopes"]
    };

    const email =
      Office?.context?.mailbox?.userProfile?.emailAddress?.toLowerCase();
    if (email) {
      options.loginHint = email;
    }

    this.msalService
      .loginRedirect({ ...options } as RedirectRequest)
      .toPromise();
  }
}