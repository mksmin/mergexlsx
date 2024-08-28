"""
Существует мастер-файл с персональными данными участников,
а так же информацией о прохождении ими курса в виде  Excel таблицы

Этот модуль создает в PostgreSQL таблицу за заданными в классе UserSt названиями столбцов
Открывает Excel файл и построчно добавляет каждое значение в БД
"""


# Импорт библиотек
import asyncio, os
import logging
import pandas as pd

# Импорт функций из библиотек
from sqlalchemy import BigInteger, String, Column, TIMESTAMP
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from sqlalchemy.ext.asyncio import AsyncAttrs, async_sessionmaker, create_async_engine
from dotenv import load_dotenv

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(dotenv_path)

post_host_token = os.getenv('TOKENSQL')
engine = create_async_engine(url=post_host_token, echo=True)
async_session = async_sessionmaker(engine)


class Base(AsyncAttrs, DeclarativeBase):
    pass


class UserSt(Base):
    __tablename__ = 'studocehexcel'

    id: Mapped[int] = mapped_column(primary_key=True)
    user_id = mapped_column(BigInteger)
    name: Mapped[str] = mapped_column(String(300))
    compet: Mapped[str] = mapped_column(String(200))
    placeofstudy: Mapped[str] = mapped_column(String(500))
    categorypos: Mapped[str] = mapped_column(String(300))
    partner: Mapped[str] = mapped_column(String(300))
    course: Mapped[int] = mapped_column()
    speciality: Mapped[str] = mapped_column(String(300))
    leveleducation: Mapped[str] = mapped_column(String(300))
    direction: Mapped[str] = mapped_column(String(300))
    region: Mapped[str] = mapped_column(String(300))
    city: Mapped[str] = mapped_column(String(300))
    birthday = Column(TIMESTAMP)
    age: Mapped[int] = mapped_column()
    email: Mapped[str] = mapped_column(String(300))
    mobile: Mapped[str] = mapped_column(String(300))
    telegram: Mapped[str] = mapped_column(String(300))
    soft1: Mapped[str] = mapped_column(String(500), nullable=True)
    soft2: Mapped[str] = mapped_column(String(500), nullable=True)
    soft3: Mapped[str] = mapped_column(String(500), nullable=True)
    soft4: Mapped[str] = mapped_column(String(500), nullable=True)
    soft5: Mapped[str] = mapped_column(String(500), nullable=True)
    soft6: Mapped[str] = mapped_column(String(500), nullable=True)
    competentions: Mapped[str] = mapped_column(String(300))
    maxscore: Mapped[int] = mapped_column()
    score: Mapped[int] = mapped_column()
    perstudi: Mapped[int] = mapped_column()
    tecpred: Mapped[str] = mapped_column(String(500))
    soft: Mapped[int] = mapped_column()
    comp: Mapped[int] = mapped_column()
    TP: Mapped[int] = mapped_column()
    sum_score: Mapped[int] = mapped_column()


async def async_main():
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)


path_to_file = os.path.join(os.path.dirname(__file__), 'FilesXlsx\\Test.xlsx')


async def excelhelp():
    with pd.ExcelFile(path_to_file) as xls:
        value_from_excel = pd.read_excel(xls)
        col_index_start = len(value_from_excel.index)
        index = 0
        while index <= col_index_start:
            ve = value_from_excel.iloc[index]
            print(f'НАЧАЛ РАБОТУ С  {ve['id']}')
            async with async_session() as session:
                session.add(UserSt(
                    id=ve['id'],
                    user_id=ve['user_id'],
                    name=ve['Name'],
                    compet=ve['Compet'],
                    placeofstudy=ve['PlaceOfStudy'],
                    categorypos=ve['CategoryPoS'],
                    partner=ve['Partner'],
                    course=ve['Course'],
                    speciality=ve['Speciality'],
                    leveleducation=ve['LevelEducation'],
                    direction=ve['Direction'],
                    region=ve['Region'],
                    city=ve['City'],
                    birthday=ve['Birthday'],
                    age=ve['Age'],
                    email=ve['email'],
                    mobile=ve['mobile'],
                    telegram=ve['telegram'],
                    soft1=ve['soft1'],
                    soft2=ve['soft2'],
                    soft3=ve['soft3'],
                    soft4=ve['soft4'],
                    soft5=ve['soft5'],
                    soft6=ve['soft6'],
                    competentions=ve['Competentions'],
                    maxscore=ve['maxScore'],
                    score=ve['Score'],
                    perstudi=ve['PerStudi'],
                    tecpred=ve['TecPred'],
                    soft=ve['soft'],
                    comp=ve['comp'],
                    TP=ve['TP'],
                    sum_score=ve['Sum_score']
                ))
                await session.commit()
                index += 1


async def main():
    await async_main()
    print('Started database')
    await excelhelp()


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print('Exit')
