using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Sasha.Exceptions;
using Sasha.ExtensionMethods;
using XrmLibrary.EntityHelpers.Common;

namespace XrmLibrary.EntityHelpers.BusinessManagement
{
    /// <summary>
    /// The <c>Timezone</c> entities can be used when you write code that works in multiple timezones.
    /// This class provides mostly used common methods for <c>Timezone</c> entities.
    /// <para>
    /// For more information look at https://msdn.microsoft.com/en-us/library/gg309411(v=crm.8).aspx
    /// </para>
    /// </summary>
    public class TimezoneHelper : BaseEntityHelper
    {
        #region | Enums |

        /// <summary>
        /// <c>Language_Country</c> formatted Microsoft Locale Id List.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/ms912047(WinEmbedded.10).aspx
        /// </para>
        /// </summary>
        public enum LocaleId
        {
            Afrikaans_SouthAfrica = 1078,
            Albanian_Albania = 1052,
            Arabic_Algeria = 5121,
            Arabic_Bahrain = 15361,
            Arabic_Egypt = 3073,
            Arabic_Iraq = 2049,
            Arabic_Jordan = 11265,
            Arabic_Kuwait = 13313,
            Arabic_Lebanon = 12289,
            Arabic_Libya = 4097,
            Arabic_Morocco = 6145,
            Arabic_Oman = 8193,
            Arabic_Qatar = 16385,
            Arabic_SaudiArabia = 1025,
            Arabic_Syria = 10241,
            Arabic_Tunisia = 7169,
            Arabic_UAE = 14337,
            Arabic_Yemen = 9217,
            Armenian_Armenia = 1067,
            AzeriCyrillic_Azerbaijan = 2092,
            AzeriLatin_Azerbaijan = 1068,
            Basque_Spain = 1069,
            Belarusian_Belarus = 1059,
            Bulgarian_Bulgaria = 1026,
            Catalan_Spain = 1027,
            Chinese_HongKong = 3076,
            Chinese_Macau = 5124,
            ChineseSimplified_PRC = 2052,
            Chinese_Singapore = 4100,
            Chinese_Taiwan = 1028,
            Croatian_Croatia = 1050,
            Czech_CzechRepublic = 1029,
            Danish_Denmark = 1030,
            Divehi_Maldives = 1125,
            Dutch_Belgium = 2067,
            Dutch_Netherlands = 1043,
            English_Australia = 3081,
            English_Belize = 10249,
            English_Canada = 4105,
            English_Caribbean = 9225,
            English_Ireland = 6153,
            English_Jamaica = 8201,
            English_NewZealand = 5129,
            English_Philippines = 13321,
            English_SouthAfrica = 7177,
            English_TrinidadandTobago = 11273,
            English_UnitedKingdom = 2057,
            EnglishDefault_UnitedStates = 1033,
            English_Zimbabwe = 12297,
            Estonian_Estonia = 1061,
            Faroese_FaeroeIslands = 1080,
            Farsi_Iran = 1065,
            Finnish_Finland = 1035,
            French_Belgium = 2060,
            French_Canada = 3084,
            FrenchDefault_France = 1036,
            French_Luxembourg = 5132,
            French_PrincipalityofMonaco = 6156,
            French_Switzerland = 4108,
            Macedonian_Macedonia = 1071,
            Galician_Spain = 1110,
            Georgian_Georgia = 1079,
            German_Austria = 3079,
            GermanDefault_Germany = 1031,
            German_Liechtenstein = 5127,
            German_Luxembourg = 4103,
            German_Switzerland = 2055,
            Greek_Greece = 1032,
            Gujarati_India = 1095,
            Hebrew_Israel = 1037,
            Hindi_India = 1081,
            Hungarian_Hungary = 1038,
            Icelandic_Iceland = 1039,
            Indonesian_Indonesia = 1057,
            ItalianDefault_Italy = 1040,
            Italian_Switzerland = 2064,
            Japanese_Japan = 1041,
            Kannada_India = 1099,
            Kazakh_Kazakhstan = 1087,
            Konkani_India = 1111,
            Korean_Korea = 1042,
            Kyrgyz_Kyrgyzstan = 1088,
            Latvian_Latvia = 1062,
            Lithuanian_Lithuania = 1063,
            Malay_BruneiDarussalam = 2110,
            Malay_Malaysia = 1086,
            Marathi_India = 1102,
            Mongolian_Mongolia = 1104,
            NorwegianBokmal_Norway = 1044,
            NorwegianNynorsk_Norway = 2068,
            Polish_Poland = 1045,
            Portuguese_Brazil = 1046,
            Portuguese_Portugal = 2070,
            Punjabi_India = 1094,
            Romanian_Romania = 1048,
            Russian_Russia = 1049,
            Sanskrit_India = 1103,
            SerbianCyrillic_SerbiaAndMontenegro = 3098,
            SerbianLatin_SerbiaAndMontenegro = 2074,
            Slovak_Slovakia = 1051,
            Slovenian_Slovenia = 1060,
            Spanish_Argentina = 11274,
            Spanish_Bolivia = 16394,
            Spanish_Chile = 13322,
            Spanish_Colombia = 9226,
            Spanish_CostaRica = 5130,
            Spanish_DominicanRepublic = 7178,
            Spanish_Ecuador = 12298,
            Spanish_ElSalvador = 17418,
            Spanish_Guatemala = 4106,
            Spanish_Honduras = 18442,
            Spanish_Mexico = 2058,
            Spanish_Nicaragua = 19466,
            Spanish_Panama = 6154,
            Spanish_Paraguay = 15370,
            Spanish_Peru = 10250,
            Spanish_PuertoRico = 20490,
            Spanish_Spain = 1034,
            Spanish_Uruguay = 14346,
            Spanish_Venezuela = 8202,
            SpanishDefault_Spain = 3082,
            Swahili_Kenya = 1089,
            Swedish_Finland = 2077,
            Swedish_Sweden = 1053,
            Syriac_Syria = 1114,
            Tamil_India = 1097,
            Tatar_Tatarstan = 1092,
            Telugu_India = 1098,
            Thai_Thailand = 1054,
            Turkish_Turkey = 1055,
            Ukrainian_Ukraine = 1058,
            Urdu_Pakistan = 1056,
            UzbekCyrillic_Uzbekistan = 2115,
            UzbekLatin_Uzbekistan = 1091,
            Vietnamese_VietNam = 1066,
            Welsh_UnitedKingdom = 1106
        }

        /// <summary>
        /// Microsoft Dynamics CRM <c>Timezone</c> List.
        /// For more information look at <c>dbo.TimeZoneDefinition</c> table/view on <c>orgname_MSCRM</c> database.
        /// <para>
        /// Formatting parameters : 
        /// <c>P_</c> : GMT+ ,
        /// <c>S_</c> : GMT- ,
        /// <c>H</c> : Hour ,
        /// <c>M</c> : Minutes
        /// </para>
        /// <para>
        /// Formatting information : 
        /// <c>GMT(+/-)_Hour_Minute_StandartName_InterfaceName</c>
        /// </para>
        /// </summary>
        public enum TimezoneCode
        {
            UTC_UTC = 92,
            P_10H_CANBERRA__MELBOURNE__SYDNEY_Canberra_Melbourne_Sydney = 256,
            S_04H_SA_WESTERN_STANDARD_TIME_Georgetown_LaPaz_Manaus_SanJuan = 55,
            P_00H_GMT_STANDARD_TIME_Dublin_Edinburgh_Lisbon_London = 85,
            P_05H_EKATERINBURG_STANDARD_TIME_Ekaterinburg = 180,
            P_03H_ARABIC_STANDARD_TIME_Baghdad = 158,
            P_06H_BANGLADESH_STANDARD_TIME_Dhaka = 196,
            S_06H_CENTRAL_STANDARD_TIME_MEXICO_Guadalajara_MexicCity_Monterrey = 29,
            P_04H_ARABIAN_STANDARD_TIME_AbuDhabi_Muscat = 165,
            P_04H_ASTRAKHAN_STANDARD_TIME_Astrakhan_Ulyanovsk = 176,
            P_02H_MIDDLE_EAST_STANDARD_TIME_Beirut = 131,
            P_10H30M_LORD_HOWE_STANDARD_TIME_LordHoweIsland = 274,
            P_10H_AUS_EASTERN_STANDARD_TIME_Canberra_Melbourne_Sydney = 255,
            P_06H_OMSK_STANDARD_TIME_Omsk = 197,
            P_07H_ALTAI_STANDARD_TIME_Barnaul_GornoAltaysk = 208,
            P_10H_EAST_AUSTRALIA_STANDARD_TIME_Brisbane = 260,
            S_01H_CAPE_VERDE_STANDARD_TIME_CaboVerde = 83,
            P_11H_RUSSIA_TIME_ZONE_Chokurdakh = 279,
            P_01H_ROMANCE_STANDARD_TIME_Brussels_Copenhagen_Madrid_Paris = 105,
            P_02H_WEST_BANK_GAZA_STANDARD_TIME_Gaza_Hebron = 142,
            P_12H_NEW_ZEALAND_STANDARD_TIME_Auckland_Wellington = 290,
            P_02H_EAST_EUROPE_STANDARD_TIME_Chisinau = 115,
            P_02H_GTB_STANDARD_TIME_Athens_Bucharest = 130,
            P_11H_BOUGAINVILLE_STANDARD_TIME_Bougainville_Island = 276,
            P_13H_TONGA_STANDARD_TIME_Nukualofa = 300,
            P_10H_TASMANIA_STANDARD_TIME_Hobart = 265,
            P_04H_MAURITIUS_STANDARD_TIME_Port_Louis = 172,
            P_07H_NORTH_CENTRAL_ASIA_STANDARD_TIME_Novosibirsk = 201,
            P_11H_NORFOLK_STANDARD_TIME_Norfolk_Island = 277,
            S_09H_UTC_S09_Coordinated_Universal_Time = 9,
            S_03H30M_NEWFOUNDLAND_STANDARD_TIME_Newfoundland = 60,
            P_03H30M_IRAN_STANDARD_TIME_Tehran = 160,
            P_08H_CHINA_STANDARD_TIME_Beijing_Chongqing_Hong_Kong_Urumqi = 210,
            P_02H_JORDAN_STANDARD_TIME_Amman = 129,
            P_04H_RUSSIA_TIME_ZONE_Izhevsk_Samara = 174,
            S_06H_CENTRAL_AMERICA_STANDARD_TIME_Central_America = 33,
            P_07H_NORTH_ASIA_STANDARD_TIME_Krasnoyarsk = 207,
            S_11H_UTC_S11_Coordinated_Universal_Time = 6,
            P_01H_CENTRAL_EUROPEAN_STANDARD_TIME_Sarajevo_Skopje_Warsaw_Zagreb = 100,
            S_05H_SA_PACIFIC_STANDARD_TIME_Bogota_Lima_Quito_Rio_Branco = 45,
            S_03H_SAINT_PIERRE_STANDARD_TIME_Saint_Pierre_and_Miquelon = 72,
            P_09H_YAKUTSK_STANDARD_TIME_Yakutsk = 240,
            P_04H_CAUCASUS_STANDARD_TIME_Yerevan = 170,
            P_09H_TOKYO_STANDARD_TIME_Osaka_Sapporo_Tokyo = 235,
            P_09H_TRANSBAIKAL_STANDARD_TIME_Chita = 241,
            P_10H_WEST_PACIFIC_STANDARD_TIME_Guam_Port_Moresby = 275,
            P_12H_FIJI_STANDARD_TIME_Fiji = 285,
            P_07H_W_MONGOLIA_STANDARD_TIME_Hovd = 209,
            P_04H30M_AFGHANISTAN_STANDARD_TIME_Kabul = 175,
            S_05H_EASTERN_STANDARD_TIME_Eastern_Time_US_Canada = 35,
            P_10H_VLADIVOSTOK_STANDARD_TIME_Vladivostok = 270,
            S_01H_AZORES_STANDARD_TIME_Azores = 80,
            P_02H_JERUSALEM_STANDARD_TIME_Jerusalem = 135,
            P_01H_CENTRAL_EUROPE_STANDARD_TIME_Belgrade_Bratislava_Budapest_Ljubljana_Prague = 95,
            P_07H_SE_ASIA_STANDARD_TIME_Bangkok_Hanoi_Jakarta = 205,
            S_03H_ARGENTINA_STANDARD_TIME_Buenos_Aires = 69,
            P_08H_NORTH_ASIA_EAST_STANDARD_TIME_Irkutsk = 227,
            P_11H_MAGADAN_STANDARD_TIME_Magadan = 281,
            S_03H_SA_EASTERN_STANDARD_TIME_Cayenne_Fortaleza = 70,
            S_07H_MOUNTAIN_STANDARD_TIME_MEXICO_Chihuahua_La_Paz_Mazatlan = 12,
            P_02H_EGYPT_STANDARD_TIME_Cairo = 120,
            S_08H_PACIFIC_STANDARD_TIME_MEXICO_Baja_California = 5,
            P_11H_SAKHALIN_STANDARD_TIME_Sakhalin = 278,
            P_12H_RUSSIA_TIME_ZONE_Anadyr_PetropavlovskKamchatsky = 295,
            P_05H_WEST_ASIA_STANDARD_TIME_Ashgabat_Tashkent = 185,
            S_05H_HAITI_STANDARD_TIME_Haiti = 43,
            P_03H_EAST_AFRICA_STANDARD_TIME_Nairobi = 155,
            S_04H_ATLANTIC_STANDARD_TIME_Atlantic_Time_Canada = 50,
            S_04H_TURKS_AND_CAICOS_STANDARD_TIME_Turks_and_Caicos = 51,
            S_12H_DATELINE_STANDARD_TIME_International_Date_Line_West = 0,
            P_02H_KALININGRAD_STANDARD_TIME_Kaliningrad = 159,
            S_02H_UTC_S02_Coordinated_Universal_Time = 76,
            P_03H_TURKEY_STANDARD_TIME_Istanbul = 134,
            P_08H_WEST_AUSTRALIA_STANDARD_TIME_Perth = 225,
            P_13H_SAMOA_STANDARD_TIME_Samoa = 1,
            P_00H_MOROCCO_STANDARD_TIME_Casablanca = 84,
            P_06H_CENTRAL_ASIA_STANDARD_TIME_Astana = 195,
            P_09H30M_AUS_CENTRAL_STANDARD_TIME_Darwin = 245,
            S_04H_VENEZUELA_STANDARD_TIME_Caracas = 47,
            S_10H_HAWAIIAN_STANDARD_TIME_Hawaii = 2,
            P_09H_KOREA_STANDARD_TIME_Seoul = 230,
            P_08H_SINGAPORE_STANDARD_TIME_Kuala_Lumpur_Singapore = 215,
            S_09H30M_MARQUESAS_STANDARD_TIME_Marquesas_Islands = 8,
            P_07H_TOMSK_STANDARD_TIME_Tomsk = 211,
            S_06H_MEXICO_STANDARD_TIME_Guadalajara_MexicoCity_MonterreyOld = 30,
            P_06H30M_MYANMAR_STANDARD_TIME_Yangon = 203,
            P_03H_RUSSIAN_STANDARD_TIME_Moscow_StPetersburg_Volgograd = 145,
            P_02H_SYRIA_STANDARD_TIME_Damascus = 133,
            S_07H_US_MOUNTAIN_STANDARD_TIME_Arizona = 15,
            P_04H_GEORGIAN_STANDARD_TIME_Tbilisi = 173,
            S_09H_ALASKAN_STANDARD_TIME_Alaska = 3,
            S_03H_EAST_SOUTH_AMERICA_STANDARD_TIME_Brasilia = 65,
            P_01H_WEST_EUROPE_STANDARD_TIME_Amsterdam_Berlin_Bern_Rome_Stockholm_Vienna = 110,
            P_08H30M_NORTH_KOREA_STANDARD_TIME_Pyongyang = 229,
            S_07H_MEXICO_STANDARD_TIME_Chihuahua_LaPaz_MazatlanOld = 13,
            P_08H_ULAANBAATAR_STANDARD_TIME_Ulaanbaatar = 228,
            S_03H_GREENLAND_STANDARD_TIME_Greenland = 73,
            P_09H30M_CENTRAL_AUSTRALIA_STANDARD_TIME_Adelaide = 250,
            S_02H_MID_ATLANTIC_STANDARD_TIME_MidAtlantic = 75,
            P_12H_UTC_P12_Coordinated_Universal_Time = 284,
            P_02H_FLE_STANDARD_TIME_Helsinki_Kyiv_Riga_Sofia_Tallinn_Vilnius = 125,
            P_02H_SOUTH_AFRICA_STANDARD_TIME_Harare_Pretoria = 140,
            S_05H_US_EASTERN_STANDARD_TIME_Indiana_East = 40,
            P_01H_WEST_CENTRAL_AFRICA_STANDARD_TIME_West_Central_Africa = 113,
            S_08H_PACIFIC_STANDARD_TIME_Pacific_Time_US_Canada = 4,
            P_05H_PAKISTAN_STANDARD_TIME_Islamabad_Karachi = 184,
            P_03H_BELARUS_STANDARD_TIME_Minsk = 151,
            P_11H_CENTRAL_PACIFIC_STANDARD_TIME_Solomon_Is_New_Caledonia = 280,
            S_03H_TOCANTINS_STANDARD_TIME_Araguaina = 77,
            P_03H_ARAB_STANDARD_TIME_Kuwait_Riyadh = 150,
            S_05H_CUBA_STANDARD_TIME_Havana = 44,
            P_08H_TAIPEI_STANDARD_TIME_Taipei = 220,
            S_05H_EASTERN_STANDARD_TIME_MEXICO_Chetumal = 301,
            S_03H_MONTEVIDEO_STANDARD_TIME_Montevideo = 74,
            P_09H30M_ADELAIDE_Adelaide = 251,
            P_10H_HOBART_Hobart = 266,
            P_05H30M_SRI_LANKA_STANDARD_TIME_Sri_Jayawardenepura = 200,
            S_10H_ALEUTIAN_STANDARD_TIME_Aleutian_Islands = 7,
            P_04H_AZERBAIJAN_STANDARD_TIME_Baku = 169,
            P_05H30M_INDIA_STANDARD_TIME_Chennai_Kolkata_Mumbai_NewDelhi = 190,
            S_06H_EASTER_ISLAND_STANDARD_TIME_Easter_Island = 34,
            S_06H_CENTRAL_STANDARD_TIME_Central_Time_US_Canada = 20,
            S_08H_UTC_S08_Coordinated_Universal_Time = 11,
            S_07H_MOUNTAIN_STANDARD_TIME_Mountain_Time_US_Canada = 10,
            S_06H_CANADA_CENTRAL_STANDARD_TIME_Saskatchewan = 25,
            P_08H45M_AUS_CENTRAL_WEST_STANDARD_TIME_Eucla = 231,
            S_04H_CENTRAL_BRAZILIAN_STANDARD_TIME_Cuiaba = 58,
            S_04H_PARAGUAY_STANDARD_TIME_Asuncion = 59,
            P_01H_NAMIBIA_STANDARD_TIME_Windhoek = 141,
            S_04H_PACIFIC_SA_STANDARD_TIME_Santiago = 56,
            P_05H45_NEPAL_STANDARD_TIME_Kathmandu = 193,
            S_03H_BAHIA_STANDARD_TIME_Salvador = 71,
            P_00H_GREENWICH_STANDARD_TIME_Monrovia_Reykjavik = 90,
            P_12H45M_CHATHAM_ISLANDS_STANDARD_TIME_Chatham_Islands = 299
        }

        #endregion

        #region | Constructors |

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service"><see cref="IOrganizationService"/></param>
        public TimezoneHelper(IOrganizationService service) : base(service)
        {
            this.SdkHelpUrl = "https://msdn.microsoft.com/en-us/library/gg309411(v=crm.8).aspx";
            this.EntityName = "timezone";
        }

        #endregion

        #region | Public Methods |

        /// <summary>
        /// Retrieve all the time zone definitions for the specified locale and to return only the <c>Display Name</c> attribute.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.getalltimezoneswithdisplaynamerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="localeId">
        /// <see cref="LocaleId"/> info.
        /// For more information look at https://msdn.microsoft.com/en-us/library/ms912047(WinEmbedded.10).aspx
        /// </param>
        /// <returns>
        /// <see cref="EntityCollection"/> for <c>Timezone</c> data
        /// </returns>
        public EntityCollection GetAllTimezoneData(LocaleId localeId)
        {
            GetAllTimeZonesWithDisplayNameRequest request = new GetAllTimeZonesWithDisplayNameRequest()
            {
                LocaleId = (int)localeId
            };

            var serviceResponse = (GetAllTimeZonesWithDisplayNameResponse)this.OrganizationService.Execute(request);
            return serviceResponse.EntityCollection;
        }

        /// <summary>
        /// Retrieve <c>LocalTime</c> from <c>UtcTime</c> by <see cref="IOrganizationService"/> <c>Caller</c> 's settings.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.localtimefromutctimerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="utcTime">
        /// <see cref="DateTime"/> UtcTime
        /// </param>
        /// <returns>
        /// <see cref="DateTime"/> LocalTime
        /// </returns>
        public DateTime GetLocalTimeFromUTCTime(DateTime utcTime)
        {
            var userSettings = GetCurrentUsersSettings(null);

            LocalTimeFromUtcTimeRequest request = new LocalTimeFromUtcTimeRequest()
            {
                UtcTime = utcTime,
                TimeZoneCode = userSettings.Item2
            };

            var serviceResponse = (LocalTimeFromUtcTimeResponse)this.OrganizationService.Execute(request);
            return serviceResponse.LocalTime;
        }

        /// <summary>
        /// Retrieve <c>LocalTime</c> from <c>UtcTime</c> by specified <c>System User</c> 's settings.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.localtimefromutctimerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="utcTime">
        /// <see cref="DateTime"/> UtcTime
        /// </param>
        /// <param name="userId"><c>System User</c> Id</param>
        /// <returns>
        /// <see cref="DateTime"/> LocalTime
        /// </returns>
        public DateTime GetLocalTimeFromUTCTime(DateTime utcTime, Guid userId)
        {
            ExceptionThrow.IfGuidEmpty(userId, "userId");

            var userSettings = GetCurrentUsersSettings(userId);

            LocalTimeFromUtcTimeRequest request = new LocalTimeFromUtcTimeRequest()
            {
                UtcTime = utcTime,
                TimeZoneCode = userSettings.Item2
            };

            var serviceResponse = (LocalTimeFromUtcTimeResponse)this.OrganizationService.Execute(request);
            return serviceResponse.LocalTime;
        }

        /// <summary>
        /// Retrieve <c>LocalTime</c> from <c>UtcTime</c> by <c>TimeZoneCode</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.localtimefromutctimerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="utcTime">
        /// <see cref="DateTime"/> UtcTime
        /// </param>
        /// <param name="timezoneCode">You can use <see cref="TimezoneCode"/> data or directly provide Timezone Code.</param>
        /// <returns>
        /// <see cref="DateTime"/> LocalTime
        /// </returns>
        public DateTime GetLocalTimeFromUTCTime(DateTime utcTime, int timezoneCode)
        {
            LocalTimeFromUtcTimeRequest request = new LocalTimeFromUtcTimeRequest()
            {
                UtcTime = utcTime,
                TimeZoneCode = timezoneCode
            };

            var serviceResponse = (LocalTimeFromUtcTimeResponse)this.OrganizationService.Execute(request);
            return serviceResponse.LocalTime;
        }

        /// <summary>
        /// Retrieve <c>UtcTime</c> from <c>LocalTime</c> by <see cref="IOrganizationService"/> caller <c>System User</c> 's settings
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.utctimefromlocaltimerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="localTime">
        /// <see cref="DateTime"/> LocalTime
        /// </param>
        /// <returns>
        /// <see cref="DateTime"/> UtcTime
        /// </returns>
        public DateTime GetUTCTimeFromLocalTime(DateTime localTime)
        {
            var userSettings = GetCurrentUsersSettings(null);

            UtcTimeFromLocalTimeRequest request = new UtcTimeFromLocalTimeRequest()
            {
                LocalTime = localTime,
                TimeZoneCode = userSettings.Item2
            };

            var serviceResponse = (UtcTimeFromLocalTimeResponse)this.OrganizationService.Execute(request);
            return serviceResponse.UtcTime;
        }

        /// <summary>
        /// Retrieve <c>UtcTime</c> from <c>LocalTime</c> by specified <c>System User</c> 's settings.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.utctimefromlocaltimerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="localTime"><see cref="DateTime"/> LocalTime</param>
        /// <param name="userId">Systemuser Id</param>
        /// <returns>
        /// <see cref="DateTime"/> UtcTime
        /// </returns>
        public DateTime GetUTCTimeFromLocalTime(DateTime localTime, Guid userId)
        {
            ExceptionThrow.IfGuidEmpty(userId, "userId");

            var userSettings = GetCurrentUsersSettings(userId);

            UtcTimeFromLocalTimeRequest request = new UtcTimeFromLocalTimeRequest()
            {
                LocalTime = localTime,
                TimeZoneCode = userSettings.Item2
            };

            var serviceResponse = (UtcTimeFromLocalTimeResponse)this.OrganizationService.Execute(request);
            return serviceResponse.UtcTime;
        }

        /// <summary>
        /// Retrieve <c>LocalTime</c> from <c>UtcTime</c> by <c>TimeZoneCode</c>.
        /// <para>
        /// For more information look at https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.utctimefromlocaltimerequest(v=crm.8).aspx
        /// </para>
        /// </summary>
        /// <param name="localTime">
        /// <see cref="DateTime"/> LocalTime
        /// </param>
        /// <param name="timezoneCode">You can use <see cref="TimezoneCode"/> data or directly provide Timezone Code.</param>
        /// <returns>
        /// <see cref="DateTime"/> UtcTime
        /// </returns>
        public DateTime GetUTCTimeFromLocalTime(DateTime localTime, int timezoneCode)
        {
            UtcTimeFromLocalTimeRequest request = new UtcTimeFromLocalTimeRequest()
            {
                LocalTime = localTime,
                TimeZoneCode = timezoneCode
            };

            var serviceResponse = (UtcTimeFromLocalTimeResponse)this.OrganizationService.Execute(request);
            return serviceResponse.UtcTime;
        }

        #endregion

        #region | Private Methods |

        /// <summary>
        /// Retrieve <c>System User</c>'s current settings.
        /// Item1 : LocaleId
        /// Item2 : TimezoneCode
        /// </summary>
        /// <param name="id">
        /// Set <c>System User</c> Id if you process specific user, otherwise this parameter will be equals service caller user Id
        /// </param>
        /// <returns>
        /// Item1 : LocaleId , Item2 : TimezoneCode
        /// </returns>
        Tuple<int, int> GetCurrentUsersSettings(Guid? id)
        {
            Tuple<int, int> result = new Tuple<int, int>(0, 0);

            QueryExpression query = new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression()
            };

            if (id.HasValue && !id.Value.IsGuidEmpty())
            {
                query.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, id.Value);
            }
            else
            {
                query.Criteria.AddCondition("systemuserid", ConditionOperator.EqualUserId);
            }

            var serviceResponse = this.OrganizationService.RetrieveMultiple(query);

            if (serviceResponse != null && serviceResponse.Entities != null && serviceResponse.Entities.Count > 0)
            {
                int localeId = 0;
                int timezoneCode = 0;

                var userSettings = serviceResponse.Entities[0];
                int.TryParse(userSettings["localeid"].ToString(), out localeId);
                int.TryParse(userSettings["timezonecode"].ToString(), out timezoneCode);

                result = new Tuple<int, int>(localeId, timezoneCode);
            }

            return result;
        }

        #endregion
    }
}
