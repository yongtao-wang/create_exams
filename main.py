# -*- coding: utf-8 -*-
import config
import utils

if __name__ == '__main__':
    print 'building xml'
    utils.create_question_xml(config.db_source)
    utils.create_exam(
        xml=config.xml_db,
        single_choice=20,
        multi_choice=20,
        blanks=10
    )
