import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SharepointChoiceComponent } from './sharepoint-choice.component';

describe('SharepointChoiceComponent', () => {
  let component: SharepointChoiceComponent;
  let fixture: ComponentFixture<SharepointChoiceComponent>;

  beforeEach(() => {
    TestBed.configureTestingModule({
      declarations: [SharepointChoiceComponent]
    });
    fixture = TestBed.createComponent(SharepointChoiceComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
