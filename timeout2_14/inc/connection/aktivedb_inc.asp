<%
public strConnect_aktiveDB
function aktivedb(x)
		
		Select Case x
	    Case 1
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intranet;" 
		'Case 2
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sthaus;"  
		'Case 3
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_buying;"
		'Case 4
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inlead;"
		'Case 5
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netstrategen;"
		'Case 6
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ravnit;"
		'Case 7
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_titoonic;"
		'Case 8
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_userneeds;"
		Case 9
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_demo;"
		'case 10
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mezzo;"
		'case 11
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skybrud;"
		'case 12
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gramtech;"
		'case 13
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cybervision;"
		'case 14
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ferro;"
		'case 15
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_henton;"
		'case 16
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_worldiq;"
		'case 17
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lysta;"
		'case 18
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_execon;"
		'case 19
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_inclusive;"
		'case 20
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kringit;"
		'case 21
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kasters;"
		'case 22
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_skousen;"
		case 23
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dencker;"
		'case 24
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_firmaservice;"
		'case 25
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_webmasters;"
		'case 26
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_proveno;"
		'case 27
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oberg;"
		'case 28
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bika;"
		'case 29
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_norma;"
		'case 30
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_stejle;"
		'case 31
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_novo_qdc;"
		'case 32
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_simi;"
		'case 33
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_external;"
		'case 34
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netkoncept;"
		'case 35
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_perspektivait;"
		case 36
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansoft;"
		'case 37
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bowe;"
		'case 38
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bminds"
		'case 39
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_workpartners"
		'case 40
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_gp"
		'case 41
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_optimizer4u"
		'case 42
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_zonerne"
		'case 43
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bolsjebutikken"
		'case 44
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_maintain"
		case 45 '**** Start ***
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_start"
		'case 46
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fyn-bo;"
	    'case 47
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_herbo;"
	    'case 48
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_radius;"
	    'case 49
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mansfield;"
	    'case 50
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mhm;"
	    'case 51
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sponsorcar;"
	    'case 52
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_abt;"
	    case 53
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_infow;"
	    case 54
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_infowdemo;"
	    case 55
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_acc;"
	    'case 56
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_elkar;"
	    'case 57
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_netfabrik;"
	    'case 58
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ingelise;"
	    'case 59
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_pcm;"
	    'case 60
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ib;"
	    'case 61
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_userminds;"
	    case 62
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_assurator;"
	    'case 63
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bsn_se;"
	    'case 64
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_syncronic;"
	    case 65
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_viatech;"
	    'case 66
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kits;"
	    'case 67
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_jt;"
	    'case 68
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_hh;"
	    'case 69
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_pcmanden;"
	    'case 70
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_accendo;"
	    'case 71
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_dreist;"
	    'case 72
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_abu;"
	    'case 73
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lna;"
	    'case 74
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_upsite;"
	    'case 75
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_spritelab;"
	    'case 76
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sec-it;"
	    'case 77
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_Q2con;"
	    case 78
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fe;"
	    'case 79
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_timeout;"
	    'case 80
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_rosenloeve;"
	    case 81
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_jm;"
	    'case 82
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_printcon;"
		'case 83
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_weboteket;"
		'case 84
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_madsen;"
		'case 85
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_raahede;"
		'case 86
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_optionone;"
		case 87
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fk;"
		case 88
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_immenso;"
		case 89
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_essens;"
		'case 90
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_svea;"
		case 91
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=93.161.131.214; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cst;"
		'case 92
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ets-track;"
		'case 93
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_berens;"
		case 94
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wowern;"
		case 95
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fkfk;"
		'case 96
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sd;"
		'case 97
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bd;"
		'case 98
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_2l;"
		'case 99
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_symbion;"
		'case 100
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_tooltest;"
		'case 101
	    'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_af;"
		case 102
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_catitest;"
        'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cma;"
		case 103
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_jep;"
		case 104
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_rage;"
		case 105
	    strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_enho;"
		'case 106
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_is;"
		case 107
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi;"
		case 108
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_jttek;"
		case 109
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wwf;"
        case 110
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mmmi;"
        case 111
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_commutemedia;"
        case 112
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_testhuset;"
        case 113
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bm;"
        case 114
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_qwert;"
        case 115
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_mi;"
        case 116
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_fk_bpm;"
        case 117
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_grundfos;"
        case 118
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_unik;"
        case 119
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_synergi1;"
        case 120
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_bst;"
        case 121
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lw;"
        case 122
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wwf2;"
        case 123
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kejd_pb;"
        case 124
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kejd_pb2;"
		case 125
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_hvk_bbb;"
		
        case 126
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_no;"
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_no;"
		case 127
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_sta;"
		case 128
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_kestner;"
		case 129
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ord;"
		case 130
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_rek;"
		case 131
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_auh;"
		case 132
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ngf;"
		case 133
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_ab;"
		case 134
		strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lawaba;"
        case 135
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_intg;"
        case 136
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_hidalgo;"
	    case 137
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_biofac;"
		case 138
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_oko;"
		case 139
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_hestia;"	
        case 140
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_epi_uk;"
        case 141
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_nt;"
        case 142
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_otvvs;"
        case 143
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_tec;"
        case 144
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_zylinc;"
        case 145
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_sdutek;"
        case 146
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_glad;"
        case 147
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_lyng;"
        case 148
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_esn;"
        case 149
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_nonstop;"
        case 150
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_danwatch;"
        case 151
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_micmatic;"
        case 152
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_aalund;"
        case 153
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_akelius;"
        case 154
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_ascendis;"
        case 155
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wilke;"
        case 156
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_adra;"
        case 157
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_cc;"
        case 158
        strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_krj;"

   


        'case 132
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=localhost; Port=3306; uid=outzource;pwd=SKba200473;database=timeout_wwf_2013;"
		
		'Case 
		'strConnect_aktiveDB = "driver={MySQL ODBC 3.51 Driver};server=192.168.1.33;Port=3306; uid=outzource;pwd=SKba200473;database=timeout_margin;"
	'* Margin Media fejler
	case else
	strConnect_aktiveDB = "nogo"
	End Select


end function
%>
