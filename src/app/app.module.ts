import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { ReactiveFormsModule } from '@angular/forms';
import { HomeComponent } from "./home/home.component";
import { Home2Component } from './home2/home2.component';
import { Home3Component } from './home3/home3.component';

@NgModule({
    declarations: [
        AppComponent,
        HomeComponent,
        Home2Component,
        Home3Component
    ],
    providers: [],
    bootstrap: [AppComponent],
    imports: [
        BrowserModule,
        AppRoutingModule,
        ReactiveFormsModule
    ]
})
export class AppModule { }
