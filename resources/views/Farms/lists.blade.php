@extends('layouts.app')

@section('content')
    <div class="row">
        <table class="table table-bordered m-3 p-6">
            <thead>
            <tr>
                <th>Predio</th>
                <th>Nombre</th>
                <th>Base de Datos Incluida</th>
                <th>Acción</th>
            </tr>
            </thead>
            <tbody>
            @foreach($list AS $item)
                <tr>
                    <td>{{$item->id_predio}}</td>
                    <td>{{$item->owner}}</td>
                    <td>{{$item->database->count()}}</td>
                    <td><a><span class="btn btn-primary"></span></a></td>
                </tr>
            @endforeach
            </tbody>
        </table>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection
