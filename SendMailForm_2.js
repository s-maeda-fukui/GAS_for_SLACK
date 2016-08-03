/*
 * Googleフォーム自動返信スクリプト
 */

function sendMailForm() {

    //------------------------------------------------------------
    // 返信メール設定エリアここから
    //------------------------------------------------------------

    //subject（件名の設定）
    var subject = '七夕祭当日スタッフへのご登録完了のお知らせ';

    //body（本文）　初期値
    //body += は、body = body + "文章"で
    //文章の結合を表しています。
    // \n　は改行
    var body = '七夕祭当日スタッフにご登録いただいた皆様\n\n'
    body += 'この度は、第１１回七夕祭当日スタッフにご登録いただき誠にありがとうございます。\n';
    body += '以下の内容でご登録完了しましたのでお知らせ致します。\nご確認お願い致します。\n';
    body += '\n'

    //footer（メール最下部）の設定
    var footer = '当日の流れ等詳細につきましては後日改めてご連絡致します。\n';
    footer += 'また、やむを得ない理由でキャンセルされる場合は七夕祭開催一週間前までに下記メールアドレスまで必ずご連絡をよろしくお願い致します。\n\n';
    footer += '七夕祭の成功は皆様にかかっています！\n'
    footer += 'どうぞよろしくお願い致します。\n'
    footer += '何かご質問等ございましたら、下記メールアドレスまでお問い合わせください。\n\n';
    footer += '-- \n';
    footer += '神戸大学 六甲台学生評議会(法・経済・経営学部ゼミ幹事会)\n'
    footer += 'B.E.L. Student Council\n';
    footer += 'bel.tanabata2016@gmail.com'

    //本文追加事項、空の状態で初期化
    var add_tmp = '';

    // 入力カラム名の指定
    // 必要なデータの入っているカラム（列名）を設定
    var NAME_COL_NAME = '名前（フルネーム）';
    var NAME_DEP = '学部(神戸大学以外の方は大学名をご記入ください)';
    var MAIL_COL_NAME = 'メールアドレス';
    var NUM_PHONE = '電話番号';
    var PART_1 = '担当したい役割(第１希望)';
    var PART_2 = '担当したい役割(第２希望)';
    var PART_3 = '担当したい役割(第３希望)';
    var CHARGE_TIME = '参加可能時間';

    //メール除外カラム（配列で処理）
    //その行（対象者）に返信をしたか、いつ返信したかのデータを格納
    var EXCLUDE_COLS = ['ステータス','対応日時'];
    // 対応日時格納変数
    var stmp ;

    // メール送信元（自分or団体のアドレス）
    var admin = 'bel.tanabata2016@gmail.com';

    //------------------------------------------------------------
    // 返信メール設定エリアここまで
    //------------------------------------------------------------
    // スプレッドシートの操作
    //アクティブなブックを取得
    var book = SpreadsheetApp.getActive();
    //名前（DB）からスプレッドシートを取得
    var sh   = book.getSheetByName('DB');
    //シートの最終行を取得
    var rows = sh.getLastRow();
    //シートの最終列を取得
    var cols = sh.getLastColumn();
    //データが入力されている範囲を取得
    var rg   = sh.getDataRange();

    //返信先アドレスを代入する変数
    var to = '';

    /* メール件名・本文作成と送信先メールアドレス取得 */
    //jは列指定（入力のある）最終列まで回す。j++はj=j+1の意味
    for ( var j = 1; j <= cols-2; j++ ) {
        // カラム名（最上位行を参照）を取得
        var col_name  = rg.getCell(1, j).getValue();
        // rows：最終行が格納されている
        // 入力値を取得
        var col_value = rg.getCell(rows, j).getValue();

        // メール用変換
        if ( col_name === NAME_COL_NAME ) {
            //格納値が”名前”のとき
            body += '【'+col_name+'】\n　';
            body += col_value+' 様\n';

        }else if ( col_name === MAIL_COL_NAME ) {
            //格納値が”メールアドレス”のとき
            to = col_value;

        }else if( col_name === NAME_DEP ){
            //格納値が'学部(神戸大学以外の方は大学名をご記入ください)
            body += '【'+col_name+'】\n　';
            body += col_value+'\n';
        }else if( col_name === NUM_PHONE ){
            //格納値が'電話番号'
            body += '【'+col_name+'】\n　';
            body += col_value+'\n';
        }else if( col_name === PART_1 ){
            //格納値が'担当したい役割(第１希望)'
            body += '【'+col_name+'】\n　';
            body += col_value+'\n';
        }else if( col_name === PART_2 ){
            //格納値が'担当したい役割(第２希望)'
            body += '【'+col_name+'】\n　';
            body += col_value+'\n';
        }else if( col_name === PART_3 ){
            //格納値が'担当したい役割(第３希望)'
            body += '【'+col_name+'】\n　';
            body += col_value+'\n';
        }else if( col_name === CHARGE_TIME ){
            //格納値が'参加可能時間'
            add_tmp += '【'+col_name+'】\n　';
            add_tmp += col_value+'\n';
        }else{
          //DO NOTHING
        }//if メール用変換

        // 日付フォーマットの変換
        // 他にも変換したいカラムがある場合はこのif分をコピーしてカラム名・日付フォーマットを変更する
        if ( col_name === 'タイムスタンプ' ) {
            col_value = Utilities.formatDate(col_value, 'JST', 'yyyy-MM-dd HH:mm:ss');
            //タイムスタンプをコピー
            stmp = col_value ;
        }

        // メール送信除外カラム
        //EXCLUDE_COLS.length should be 2
        if ( EXCLUDE_COLS.length > 0 ) {
            //is_exclude変数をfalseで初期化
            var is_exclude = false;
            //EXCLUDE_COLS.length should be 2
            for ( var k = 0; k < EXCLUDE_COLS.length; k++ ) {
                if ( col_name === EXCLUDE_COLS[k] ) {
                    // === は同値演算子
                    is_exclude = true;
                    break;
                }//if
            }//for

            if ( is_exclude ) {
                // 除外カラムなのでスキップ
                continue;
            }//if
        }//if
        //メール送信除外カラム部分（上）は、このコードでは不要なのですが
        //汎用性を持たせるために記述しております。

    }//for_j

    //最終連結
    body += add_tmp;
    body += '\n';
    //フッターの連結
    body += footer;

    // メール送信
    if ( to ) {
        //フォーム登録者に送信
        MailApp.sendEmail(to, subject, body);
        //シートに返信済みであること、返信日時を記入
        sh.getRange(rows, cols-1).setValue('DONE');
        sh.getRange(rows, cols).setValue(stmp);

        //管理者にも同様のメールを送信
        //MailApp.sendEmail(admin, subject, body);

    }else{
        //ほぼ不要だが、エラー処理のため設定
        MailApp.sendEmail(admin, '【失敗】Googleフォームにメールアドレスが指定されていません', body);
    }//if

    // 連続で送るとエラーになるので1秒スリープ
    Utilities.sleep(1000);


}//sendMail
