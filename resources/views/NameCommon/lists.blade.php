@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-12">
                <div class="card">
                    <div class="card-header bg-dark text-white">{{ __('Lista de Nombre comunes Nuevas') }}</div>

                    <div class="card-body">
        <div class="m-2"><a class="btn btn-success" href="{{route('get-scientific')}}" >Registar Nueva Nombre Común</a></div>
        <table class="table table-bordered ">
            <thead>
            <tr>
                <th>id</th>
                <th>Nombre Científico</th>
                <th>Nombre</th>
                <th>Acción</th>
            </tr>
            </thead>
            <tbody>
            @foreach($list AS $key =>$item)
            <tr>
                <td>{{$key+1}}</td>
                <td>{{$item->scientific->name}}</td>
                <td>{{$item->name}}</td>
                <td><a class="btn btn-primary btn-sm" href="{{route('edit-common',$item->id)}}"><i class="fa fa-pencil" aria-hidden="true"></i></a></td>
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
