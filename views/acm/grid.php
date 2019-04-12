<?php

/* @var $this yii\web\View */
/* @var $form yii\bootstrap\ActiveForm */
/* @var $model app\models\LoginForm */

use yii\helpers\Html;
use yii\bootstrap\ActiveForm;
use kartik\grid\GridView;
use yii\helpers\ArrayHelper;

$this->title = 'Acm';
$this->params['breadcrumbs'][] = $this->title;
?>
<div class="site-login">
    <h1><?= Html::encode($this->title) ?></h1>

    <p>Info:</p>
   <?php echo GridView::widget([
    'dataProvider'=>$dataProvider,
    'filterModel'=>$searchModel,
    'showPageSummary'=>true,
    'pjax'=>true,
    'striped'=>false,
    'hover'=>true,
    'panel'=>['type'=>'primary', 'heading'=>'Grid Grouping Example'],
    'columns'=>[
        ['class'=>'kartik\grid\SerialColumn'],
        [
            'attribute'=>'supplier_id', 
            'width'=>'310px',
            'value'=>function ($model, $key, $index, $widget) { 
                return $model->supplier->company_name;
            },
            'filterType'=>GridView::FILTER_SELECT2,
            'filter'=> ArrayHelper::map(Suppliers::find()->orderBy('company_name')->asArray()->all(), 'id', 'company_name'), 
            'filterWidgetOptions'=>[
                'pluginOptions'=>['allowClear'=>true],
            ],
            'filterInputOptions'=>['placeholder'=>'Any supplier'],
            'group'=>true,  // enable grouping
            'groupFooter'=>function ($model, $key, $index, $widget) { // Closure method
                return [
                    'mergeColumns'=>[[1,3]], // columns to merge in summary
                    'content'=>[             // content to show in each summary cell
                        1=>'Summary (' . $model->supplier->company_name . ')',
                        4=>GridView::F_AVG,
                        5=>GridView::F_SUM,
                        6=>GridView::F_SUM,
                    ],
                    'contentFormats'=>[      // content reformatting for each summary cell
                        4=>['format'=>'number', 'decimals'=>2],
                        5=>['format'=>'number', 'decimals'=>0],
                        6=>['format'=>'number', 'decimals'=>2],
                    ],
                    'contentOptions'=>[      // content html attributes for each summary cell
                        1=>['style'=>'font-variant:small-caps'],
                        4=>['style'=>'text-align:right'],
                        5=>['style'=>'text-align:right'],
                        6=>['style'=>'text-align:right'],
                    ],
                    // html attributes for group summary row
                    'options'=>['class'=>'danger','style'=>'font-weight:bold;']
                ];
            }
        ],
        [
            'attribute'=>'category_id', 
            'width'=>'250px',
            'value'=>function ($model, $key, $index, $widget) { 
                return $model->category->category_name;
            },
            'filterType'=>GridView::FILTER_SELECT2,
            'filter'=>ArrayHelper::map(Categories::find()->orderBy('category_name')->asArray()->all(), 'id', 'category_name'), 
            'filterWidgetOptions'=>[
                'pluginOptions'=>['allowClear'=>true],
            ],
            'filterInputOptions'=>['placeholder'=>'Any category'],
            'group'=>true,  // enable grouping
            'subGroupOf'=>1, // supplier column index is the parent group,
            'groupFooter'=>function ($model, $key, $index, $widget) { // Closure method
                return [
                    'mergeColumns'=>[[2, 3]], // columns to merge in summary
                    'content'=>[              // content to show in each summary cell
                        2=>'Summary (' . $model->category->category_name . ')',
                        4=>GridView::F_AVG,
                        5=>GridView::F_SUM,
                        6=>GridView::F_SUM,
                    ],
                    'contentFormats'=>[      // content reformatting for each summary cell
                        4=>['format'=>'number', 'decimals'=>2],
                        5=>['format'=>'number', 'decimals'=>0],
                        6=>['format'=>'number', 'decimals'=>2],
                    ],
                    'contentOptions'=>[      // content html attributes for each summary cell
                        4=>['style'=>'text-align:right'],
                        5=>['style'=>'text-align:right'],
                        6=>['style'=>'text-align:right'],
                    ],
                    // html attributes for group summary row
                    'options'=>['class'=>'success','style'=>'font-weight:bold;']
                ];
            },
        ],
        [
            'attribute'=>'product_name',
            'pageSummary'=>'Page Summary',
            'pageSummaryOptions'=>['class'=>'text-right text-warning'],
        ],
        [
            'attribute'=>'unit_price',
            'width'=>'150px',
            'hAlign'=>'right',
            'format'=>['decimal', 2],
            'pageSummary'=>true,
            'pageSummaryFunc'=>GridView::F_AVG
        ],
        [
            'attribute'=>'units_in_stock',
            'width'=>'150px',
            'hAlign'=>'right',
            'format'=>['decimal', 0],
            'pageSummary'=>true
        ],
        [
            'class'=>'kartik\grid\FormulaColumn',
            'header'=>'Amount In Stock',
            'value'=>function ($model, $key, $index, $widget) { 
                $p = compact('model', 'key', 'index');
                return $widget->col(4, $p) * $widget->col(5, $p);
            },
            'mergeHeader'=>true,
            'width'=>'150px',
            'hAlign'=>'right',
            'format'=>['decimal', 2],
            'pageSummary'=>true
        ],
    ],
]);
?>
</div>
