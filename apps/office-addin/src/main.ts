// eslint-disable-next-line @typescript-eslint/triple-slash-reference
/// <reference path="../../../node_modules/@types/office-js/index.d.ts" />
// eslint-disable-next-line @typescript-eslint/triple-slash-reference
/// <reference path="../../../node_modules/@types/office-runtime/index.d.ts" />

import { enableProdMode, NgModuleRef } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

declare const Office: any;

// Reference: https://www.npmjs.com/package/@microsoft/office-js
// https://docs.microsoft.com/en-us/office/dev/add-ins/develop/register-sso-add-in-aad-v2

let bootstapped = false;

if (Office !== undefined) {
  Office.initialize = () => {
    Office.onReady()
      .then(
        (_info: { host: Office.HostType; platform: Office.PlatformType }) => {
          if (!bootstapped) {
            console.log('Office is ready', _info);
            bootstapped = true;
            return bootStrapAngular();
          } else {
            return;
          }
        }
      )
      .catch((err: any) => {
        console.warn('Angular not bootstrapped. Error: \n', err);
      });
  };

  Office.initialize(0); // Start
} else {
  console.log('Bootstrapping Angular, without Office JS');

  bootStrapAngular();
}

function bootStrapAngular(): Promise<void | NgModuleRef<AppModule>> {
  return platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .then(() => {
      if ('serviceWorker' in navigator && environment.production) {
        navigator.serviceWorker.register('/ngsw-worker.js');
      }
    })
    .catch((error) => {
      console.log('Error bootstrapping Office Angular app: ', error.message);
      
    });
}