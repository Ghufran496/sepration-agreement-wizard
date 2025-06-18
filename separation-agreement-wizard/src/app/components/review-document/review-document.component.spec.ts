import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ReviewDocumentComponent } from './review-document.component';

describe('ReviewDocumentComponent', () => {
  let component: ReviewDocumentComponent;
  let fixture: ComponentFixture<ReviewDocumentComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ReviewDocumentComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ReviewDocumentComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
