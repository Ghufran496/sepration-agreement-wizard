import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router } from '@angular/router';
import { ButtonModule } from 'primeng/button';
import { ProgressSpinnerModule } from 'primeng/progressspinner';
import { WizardService } from '../../services/wizard.service';
import { PartyInfo } from '../../models/party-info';
import { Clause } from '../../models/clause';
import { finalize, take } from 'rxjs/operators';

@Component({
  selector: 'app-review-document',
  standalone: true,
  imports: [
    CommonModule,
    ButtonModule,
    ProgressSpinnerModule
  ],
  templateUrl: './review-document.component.html',
  styleUrl: './review-document.component.scss'
})
export class ReviewDocumentComponent implements OnInit {
  partyInfo: PartyInfo = {
    party1FirstName: '',
    party1LastName: '',
    party2FirstName: '',
    party2LastName: '',
    children: []
  };
  selectedClauses: Clause[] = [];
  isGenerating: boolean = false;
  loading: boolean = true;

  constructor(
    private wizardService: WizardService,
    private router: Router
  ) { }

  ngOnInit(): void {
    this.loadData();
  }

  loadData(): void {
    this.loading = true;
    console.log('Loading review document data...');
    
    this.wizardService.getPartyInfo()
      .pipe(take(1))
      .subscribe({
        next: (partyInfo) => {
          console.log('Party info loaded:', partyInfo);
          this.partyInfo = partyInfo;
          this.loadSelectedClauses();
        },
        error: (error) => {
          console.error('Error loading party info:', error);
          this.loading = false;
        }
      });
  }
  
  loadSelectedClauses(): void {
    this.wizardService.getOrderedSelectedClauses()
      .pipe(
        take(1),
        finalize(() => {
          this.loading = false;
          console.log('Data loading complete');
        })
      )
      .subscribe({
        next: (clauses) => {
          console.log('Selected clauses loaded:', clauses);
          this.selectedClauses = clauses;
        },
        error: (error) => {
          console.error('Error loading selected clauses:', error);
        }
      });
  }

  formatClauseText(text: string): string {
    if (!text || !this.partyInfo) return text;
    
    let formattedText = text;
    
    formattedText = formattedText.replace(/\*PARTY 1\*/g, 
      `${this.partyInfo.party1FirstName} ${this.partyInfo.party1LastName}`);
    formattedText = formattedText.replace(/\*PARTY 2\*/g, 
      `${this.partyInfo.party2FirstName} ${this.partyInfo.party2LastName}`);
    
    return formattedText;
  }

  getPartyName(partyId: string | undefined): string {
    if (!partyId) return '';
    
    if (partyId === 'party1') {
      return `${this.partyInfo.party1FirstName} ${this.partyInfo.party1LastName}`;
    } else if (partyId === 'party2') {
      return `${this.partyInfo.party2FirstName} ${this.partyInfo.party2LastName}`;
    }
    return '';
  }

  hasSupport(): boolean {
    return !!(this.partyInfo.spousalSupportPayor || this.partyInfo.childSupportPayor);
  }

  hasDates(): boolean {
    return !!(this.partyInfo.marriedDate || this.partyInfo.cohabitationDate || this.partyInfo.separationDate);
  }

  onBack(): void {
    this.router.navigate(['/clause-selection']);
  }

  onGenerateDocument(): void {
    this.isGenerating = true;
    
    this.wizardService.generateDocument().subscribe(
      (blob: Blob) => {
        this.isGenerating = false;
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `SeparationAgreement_${new Date().toISOString().slice(0, 10)}.docx`;
        document.body.appendChild(a);
        a.click();
        
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      },
      (error) => {
        this.isGenerating = false;
        console.error('Error generating document:', error);
        alert('Error generating document. Please try again later.');
      }
    );
  }
}
