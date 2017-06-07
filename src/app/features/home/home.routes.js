routes.$inject = ['$stateProvider'];

export default function routes($stateProvider) {
    $stateProvider
        .state('home', {
            url: '/',
            template: require('./home.html'),
            resolve: {
                dataSets: ['Data', '$http', function (Data, $http) {
                    return Data.getMany('dataSets', {
                        paging: false,
                        translate: true,
                        fields: 'id,name,uuid,displayName,periodType,dataEntryForm[htmlCode],organisationUnits[id,name,displayName]'
                    });
                }]
            },
            controller: 'HomeController',
            controllerAs: 'home'
        });
}