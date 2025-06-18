import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, FormArray, ReactiveFormsModule, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { DropdownModule } from 'primeng/dropdown';
import { CalendarModule } from 'primeng/calendar';
import { WizardService } from '../../services/wizard.service';
import { PartyInfo } from '../../models/party-info';

@Component({
  selector: 'app-party-info',
  standalone: true,
  imports: [
    CommonModule,
    ReactiveFormsModule,
    ButtonModule,
    InputTextModule,
    DropdownModule,
    CalendarModule
  ],
  templateUrl: './party-info.component.html',
  styleUrl: './party-info.component.scss'
})
export class PartyInfoComponent implements OnInit {
  partyInfoForm!: FormGroup;
  
  roleOptions = [
    { label: 'Mother', value: 'Mother' },
    { label: 'Father', value: 'Father' },
    { label: 'Wife', value: 'Wife' },
    { label: 'Husband', value: 'Husband' }
  ];
  
  partyOptions = [
    { label: 'Party 1', value: 'party1' },
    { label: 'Party 2', value: 'party2' }
  ];

  constructor(
    private fb: FormBuilder,
    private wizardService: WizardService,
    private router: Router
  ) { }

  ngOnInit(): void {
    this.initForm();
    this.loadPartyInfo();
  }

  initForm(): void {
    this.partyInfoForm = this.fb.group({
      party1FirstName: ['', Validators.required],
      party1MiddleName: [''],
      party1LastName: ['', Validators.required],
      party1Role: [''],
      party2FirstName: ['', Validators.required],
      party2MiddleName: [''],
      party2LastName: ['', Validators.required],
      party2Role: [''],
      spousalSupportPayor: [''],
      spousalSupportRecipient: [''],
      childSupportPayor: [''],
      childSupportRecipient: [''],
      marriedDate: [null],
      cohabitationDate: [null],
      separationDate: [null],
      children: this.fb.array([])
    });
  }

  loadPartyInfo(): void {
    this.wizardService.getPartyInfo().subscribe(info => {
      if (info && info.party1FirstName) {
        this.partyInfoForm.patchValue(info);
        
        while (this.children.length) {
          this.children.removeAt(0);
        }
        
        if (info.children && info.children.length > 0) {
          info.children.forEach(child => {
            this.addChild(child);
          });
        }
      }
    });
  }

  addChild(childData?: any): void {
    const childForm = this.fb.group({
      firstName: [childData?.firstName || '', Validators.required],
      middleName: [childData?.middleName || ''],
      lastName: [childData?.lastName || '', Validators.required],
      birthdate: [childData?.birthdate || null]
    });
    
    this.children.push(childForm);
  }

  get children(): FormArray {
    return this.partyInfoForm.get('children') as FormArray;
  }

  onNext(): void {
    if (this.partyInfoForm.valid) {
      this.wizardService.updatePartyInfo(this.partyInfoForm.value);
      this.router.navigate(['/clause-selection']);
    } else {
      Object.keys(this.partyInfoForm.controls).forEach(key => {
        const control = this.partyInfoForm.get(key);
        control?.markAsTouched();
      });
      
      if (this.children.length > 0) {
        this.children.controls.forEach(childGroup => {
          Object.keys(childGroup.value).forEach(key => {
            const control = childGroup.get(key);
            control?.markAsTouched();
          });
        });
      }
    }
  }

  removeChild(index: number): void {
    this.children.removeAt(index);
  }
}
