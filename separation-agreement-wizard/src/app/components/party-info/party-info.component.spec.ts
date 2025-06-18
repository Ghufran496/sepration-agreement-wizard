import { ComponentFixture, TestBed } from '@angular/core/testing';

import { PartyInfoComponent } from './party-info.component';

describe('PartyInfoComponent', () => {
  let component: PartyInfoComponent;
  let fixture: ComponentFixture<PartyInfoComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [PartyInfoComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(PartyInfoComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
