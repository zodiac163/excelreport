<?php

return [
    'sourcePath' => dirname(__DIR__),
    'messagePath' => __DIR__,
    'languages' => ['ru'],
    'translator' => 'Yii::t',
    'sort' => true,
    'overwrite' => false,
    'removeUnused' => false,
    'markUnused' => true,
    'except' => [
        '.svn',
        '.git',
        '.gitignore',
        '.gitkeep',
        '.hgignore',
        '.hgkeep',
        '/messages',
        '/BaseYii.php',
    ],
    'only' => ['*.php',],
    'format' => 'php'
];