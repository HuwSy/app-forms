import { ComponentFixture, TestBed } from '@angular/core/testing';

import { HelloWorldWebPartComponent } from './hello-world-web-part.component';

describe('HelloWorldWebPartComponent', () => {
  let component: HelloWorldWebPartComponent;
  let fixture: ComponentFixture<HelloWorldWebPartComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ HelloWorldWebPartComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(HelloWorldWebPartComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
