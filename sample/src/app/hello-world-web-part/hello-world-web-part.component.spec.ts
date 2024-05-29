import { ComponentFixture, TestBed } from '@angular/core/testing';

import { HelloWorldWebPartComponent } from './hello-world-web-part.component';

describe('HelloWorldWebPartComponent', () => {
  let component: HelloWorldWebPartComponent;
  let fixture: ComponentFixture<HelloWorldWebPartComponent>;

  beforeEach(async () => {
    TestBed.configureTestingModule({
      declarations: [ HelloWorldWebPartComponent ]
    });
    fixture = TestBed.createComponent(HelloWorldWebPartComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
