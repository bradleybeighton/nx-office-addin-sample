import { NgModule } from '@angular/core';
import { RouterModule } from '@angular/router';
import { AuthGuard } from './auth.guard';
import { HomeComponent } from './home/home.component';
import { LoginComponent } from './login/login.component';

@NgModule({
  imports: [
    RouterModule.forRoot(
      [
        {
          path: 'home',
          component: HomeComponent,
          canActivate: [AuthGuard]
        },
        {
          path: 'login',
          component: LoginComponent
        },
        { path: '', redirectTo: 'home', pathMatch: 'full' }
      ],
      {
        scrollPositionRestoration: 'enabled',
        relativeLinkResolution: 'legacy',
        useHash: true
      }
    )
  ],
  exports: [RouterModule]
})
export class AppRoutingModule {}