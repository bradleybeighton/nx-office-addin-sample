
import { Injectable } from '@angular/core';
import {
  CanActivate,
  RouterStateSnapshot,
  Router,
  ActivatedRouteSnapshot
} from '@angular/router';
import { JwtHelperService } from '@auth0/angular-jwt';

import * as store from 'store2';

const helper = new JwtHelperService();

@Injectable({
  providedIn: 'root'
})
export class AuthGuard implements CanActivate {
  constructor(private router: Router) {}

  async canActivate(
    next: ActivatedRouteSnapshot,
    state: RouterStateSnapshot
  ): Promise<boolean> {
    return await this.checkLogin();
  }

  /**
   * check that there is a access-token
   * if there isn't navigate to the accounts url to verify
   */
  async checkLogin() {
    const token = store.get('access-token');
    if (token) {
      // need to check that the token is still valid
      const isExpired = helper.isTokenExpired(token);
      if (isExpired) {
        return false;
      } 
    } else {
      this.router.navigate(['login']);
    }

    return false;
  }
}