import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ClauseSelectionComponent } from './clause-selection.component';

describe('ClauseSelectionComponent', () => {
  let component: ClauseSelectionComponent;
  let fixture: ComponentFixture<ClauseSelectionComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ClauseSelectionComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ClauseSelectionComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
