import { Routes } from '@angular/router';
import { PartyInfoComponent } from './components/party-info/party-info.component';
import { ClauseSelectionComponent } from './components/clause-selection/clause-selection.component';
import { ReviewDocumentComponent } from './components/review-document/review-document.component';

export const routes: Routes = [
  { path: '', redirectTo: 'party-info', pathMatch: 'full' },
  { path: 'party-info', component: PartyInfoComponent },
  { path: 'clause-selection', component: ClauseSelectionComponent },
  { path: 'review-document', component: ReviewDocumentComponent },
  { path: '**', redirectTo: 'party-info' }
];
