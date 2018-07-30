<p align="center">
    <a href="https://custom-it.ru" target="_blank">
        <img src="https://avatars1.githubusercontent.com/u/31646762?s=200&v=4" height="100px">
    </a>
    <h1 align="center">Yii2 ExcelReport Extension</h1>
    <br>
</p>


An extension for generate excel file from GridView content. When used together in a GridView, saves the results of filtering and sorting in a file. What you see in the GridView, and it will be imported into a file.

[![Latest Stable Version](https://poser.pugx.org/custom-it/yii2-excel-report/v/stable.svg)](https://packagist.org/packages/custom-it/yii2-excel-report)
[![Total Downloads](https://poser.pugx.org/custom-it/yii2-excel-report/downloads.svg)](https://packagist.org/packages/custom-it/yii2-excel-report)

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer require --prefer-dist custom-it/yii2-excel-report
```

or add

```
"custom-it/yii2-excel-report": "*"
```

to the require section of your `composer.json` file.


Configuration
-------------
Before using the module, configure the [queues](https://raw.githubusercontent.com/yiisoft/yii2-queue/master/docs/guide/README.md)


Usage
-----

Once the extension is installed, simply use it in your code by  :

```php
$gridColumns = [
    ['class' => 'yii\grid\SerialColumn'],
    'id',
    'name',
    'date',
    'post',
    ['class' => 'yii\grid\ActionColumn'],
];

// Render widget
echo \customit\excelreport\ExcelReport::widget([
    'columns' => $gridColumns,
    'searchClass' => get_class($searchModel),
    'searchMethod' => 'search', // The method 'search' is used by default. Fill this property if another method is used
]);

// Can be used with or without a GridView
echo GridView::widget([
    'dataProvider' => $dataProvider,
    'filterModel' => $searchModel,
    'columns' => $gridColumns
]);
```
