<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Excel,Input,File;
use App\User;

class ExcelController extends Controller
{
    //
    public function getExport(){
        return view('excel.export');
    }
    public function export(){
        $export = User::select('id','name','email','password')->get()->toArray();
        $export = [1 => 'ádas'];
        Excel::create('export_data', function($excel) use ($export){
        
            $excel->sheet('Sheet 1', function($sheet) use ($export){
                $sheet->fromArray($export);
            });
        })->export('xlsx');
    }
}
