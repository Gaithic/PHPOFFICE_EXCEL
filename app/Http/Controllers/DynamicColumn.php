<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

class DynamicColumn extends Controller
{
    public function excel()
    {
        $count = 1;
        $row = range('A', 'Z');
        foreach ($row as $elements) {
            // $excel_row = $elements;
            // $col = $elements.$count;
            $count++;
            dump('rows');
            dump($elements++);
            // dump($excel_row);
            // dump(($elements));
            // dump('column');
            // dump($col);
            // dump('count');
            // dump($count);
        }
        
    }
}
