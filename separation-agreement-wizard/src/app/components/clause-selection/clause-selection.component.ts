import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { ButtonModule } from 'primeng/button';
import { InputTextModule } from 'primeng/inputtext';
import { TreeModule } from 'primeng/tree';
import { TableModule } from 'primeng/table';
import { TreeNode } from 'primeng/api';
import { WizardService } from '../../services/wizard.service';
import { Clause } from '../../models/clause';
import { finalize } from 'rxjs/operators';

@Component({
  selector: 'app-clause-selection',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    ButtonModule,
    InputTextModule,
    TreeModule,
    TableModule
  ],
  templateUrl: './clause-selection.component.html',
  styleUrl: './clause-selection.component.scss'
})
export class ClauseSelectionComponent implements OnInit {
  clauses: TreeNode[] = [];
  selectedTreeNodes: TreeNode[] = [];
  selectedClauses: Clause[] = [];
  searchText: string = '';
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

    this.wizardService.getOrderedSelectedClauses().subscribe(clauses => {
      this.selectedClauses = clauses;
    });

    this.wizardService.ensureClausesLoaded().pipe(
      finalize(() => this.loading = false)
    ).subscribe(loaded => {
      if (loaded) {
        this.loadClauses();
        this.updateSelectedTreeNodes();
      } else {
        console.error('Failed to load clauses');
      }
    });
  }

  loadClauses(): void {
    const allClauses = this.wizardService.getAllClauses();
    if (!allClauses || allClauses.length === 0) {
      console.warn('No clauses available');
      return;
    }

    const categories = this.wizardService.getCategories();

    this.clauses = categories.map(category => {
      const categoryNode: TreeNode = {
        label: category,
        data: { category },
        expandedIcon: 'pi pi-folder-open',
        collapsedIcon: 'pi pi-folder',
        children: allClauses
          .filter(clause => clause.category === category)
          .map(clause => ({
            label: clause.label && clause.label.length > 50 ? clause.label.substring(0, 50) + '...' : clause.label,
            data: clause,
            leaf: true,
            key: clause.id.toString()
          }))
      };
      return categoryNode;
    });
  }

  updateSelectedTreeNodes(): void {
    if (!this.clauses || this.clauses.length === 0) {
      return;
    }

    this.selectedTreeNodes = [];

    const selectedClauseIds = new Set(this.selectedClauses.map(clause => clause.id));

    this.clauses.forEach(categoryNode => {
      if (!categoryNode.children || categoryNode.children.length === 0) {
        return;
      }

      const childNodes = categoryNode.children;
      let allChildrenSelected = childNodes.length > 0;
      let anyChildSelected = false;

      childNodes.forEach(childNode => {
        if (childNode.data && selectedClauseIds.has(childNode.data.id)) {
          this.selectedTreeNodes.push(childNode);
          anyChildSelected = true;
        } else {
          allChildrenSelected = false;
        }
      });

      if (allChildrenSelected) {
        this.selectedTreeNodes.push(categoryNode);
      }

      if (anyChildSelected) {
        categoryNode.expanded = true;
      }
    });
  }

  onSearch(): void {
  }

  onNodeSelect(event: any): void {
    const node = event.node;

    if (node.leaf && node.data) {
      if (!this.selectedClauses.some(c => c.id === node.data.id)) {
        this.wizardService.addSelectedClause(node.data);
      }
    } else if (node.children && node.children.length > 0) {
      this.addChildClauses(node.children);
    }
  }

  addChildClauses(children: TreeNode[]): void {
    children.forEach(childNode => {
      if (childNode.leaf && childNode.data) {
        if (!this.selectedClauses.some(c => c.id === childNode.data.id)) {
          this.wizardService.addSelectedClause(childNode.data);
        }
      } else if (childNode.children && childNode.children.length > 0) {
        this.addChildClauses(childNode.children);
      }
    });
  }

  onNodeUnselect(event: any): void {
    const node = event.node;

    if (node.leaf && node.data) {
      this.wizardService.removeSelectedClause(node.data.id);
    } else if (node.children && node.children.length > 0) {
      this.removeChildClauses(node.children);
    }
  }

  removeChildClauses(children: TreeNode[]): void {
    children.forEach(childNode => {
      if (childNode.leaf && childNode.data) {
        this.wizardService.removeSelectedClause(childNode.data.id);
      } else if (childNode.children && childNode.children.length > 0) {
        this.removeChildClauses(childNode.children);
      }
    });
  }

  removeClause(clause: Clause): void {
    this.wizardService.removeSelectedClause(clause.id);
    this.updateSelectedTreeNodes();
  }

  onBack(): void {
    this.router.navigate(['/party-info']);
  }

  onNext(): void {
    this.router.navigate(['/review-document']);
  }
}
