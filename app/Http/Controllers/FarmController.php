<?php


namespace App\Http\Controllers;


use App\Imports\ImportNewExcelFarm;
use App\Models\CommonName;
use App\Models\Family;
use App\Models\Farm;
use App\Models\ForestDataBase;
use App\Models\ScientificName;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpWord\SimpleType\TblWidth;

class FarmController extends Controller
{
    public function createFamily()
    {
        return view('Families.create');
    }

    public function createCommon()
    {
        $scientifics = ScientificName::orderBy('name')->get();
        return view('NameCommon.create',compact('scientifics'));
    }

    public function createFarm()
    {
        return view('Farms.create');
    }

    public function createScientific()
    {
        $scientifics = ScientificName::orderBy('name')->get();
        return view('NameScientific.create',compact('scientifics'));
    }

    public function importStore(Request $request)
    {
        $file=$request->file();
        $import=new ImportNewExcelFarm;

        Excel::import($import,$file['data_farm']);
        return redirect()->back();
    }


    public function listsFarms()
    {
        $list=Farm::orderBy('id_predio','ASC')->get();
        return view('Farms.lists',compact('list'));
    }


    public function listsFamily()
    {
        $list=Family::orderBy('name')->get();
        return view('Families.lists',compact('list'));
    }


    public function listsCommon()
    {
        $list=CommonName::orderBy('name')->get();
        return view('NameCommon.lists',compact('list'));
    }


    public function listsScientific()
    {
        $list=ScientificName::orderBy('name')->get();
        return view('NameScientific.lists',compact('list'));
    }
}
