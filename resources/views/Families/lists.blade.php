@extends('layouts.app')

@section('content')
    <div class="row " style="width: 100%">
    <div class="col-md-12">
        <div class="mx-5"><a class="btn btn-success" >Registar Nueva Familia</a></div>
        <table class="table table-bordered m-5 p-5">
            <thead>
            <tr>
                <th>id</th>
                <th>Nombre</th>
                <th>Acción</th>
            </tr>
            </thead>
            <tbody>
            @foreach($list AS $item)
            <tr>
                <td>{{$item->id}}</td>
                <td>{{$item->name}}</td>
                <td><a class="btn btn-primary btn-sm"><i class="fa fa-pencil" aria-hidden="true"></i></a></td>
            </tr>
            @endforeach
            </tbody>
        </table>
    </div>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection
