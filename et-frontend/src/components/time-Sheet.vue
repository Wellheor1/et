<template>
  <v-container fluid>
    <v-card>
      <v-data-table :headers="headers" :items="items" :items-per-page="10">
<!--        <template v-slot:item="{ item }">-->
<!--          <template v-for="(header, i) in headers">-->
<!--            <v-edit-dialog :return-value.sync="item[headervalue]" :key="i">{{item[header.value]}}-->
<!--              <template v-slot:input>-->
<!--                <v-text-field></v-text-field>-->
<!--              </template>-->
<!--            </v-edit-dialog>-->
<!--          </template>-->
<!--        </template>-->
        <template v-slot:body="{ items, headers }">
          <tbody>
            <tr v-for="(item,idx) in items" :key="idx">
              <td v-for="(header,key) in headers" :key="key">
                <v-edit-dialog :return-value.sync="item[header.value]" @save="save" @cancel="cancel" @open="open"
                  @close="close" large cancel-text="Отменить" save-text="Сохранить" >{{item[header.value]}}
                  <template v-slot:input>
                    <v-text-field v-model="item[header.value]" label="Edit" single-line></v-text-field>
                  </template>
                </v-edit-dialog>
              </td>
            </tr>
          </tbody>
        </template>
      </v-data-table>
    </v-card>
  </v-container>
</template>

<script>
export default {
  name: 'time-Sheet',
  data () {
    return {
      headers: [
        { text: 'ФИО', value: 'name' },
        { text: 'Учётный номер', value: 'number' },
        { text: 'Должность', value: 'post' },
        { text: '1', value: 'date1' },
        { text: '2', value: 'date2' },
        { text: '3', value: 'date3' },
        { text: '4', value: 'date4' },
        { text: '5', value: 'date5' },
        { text: '6', value: 'date6' },
        { text: '7', value: 'date7' },
        { text: '8', value: 'date8' },
        { text: '9', value: 'date9' },
        { text: '10', value: 'date10' },
        { text: '11', value: 'date11' },
        { text: '12', value: 'date12' },
        { text: '13', value: 'date13' },
        { text: '14', value: 'date14' },
        { text: '15', value: 'date15' },
        { text: 'Сумма дней с 1-15', value: 'amountHalfMonths' },
        { text: '16', value: 'date16' },
        { text: '17', value: 'date17' },
        { text: '18', value: 'date18' },
        { text: '19', value: 'date19' },
        { text: '20', value: 'date20' },
        { text: '21', value: 'date21' },
        { text: '22', value: 'date22' },
        { text: '23', value: 'date23' },
        { text: '24', value: 'date24' },
        { text: '25', value: 'date25' },
        { text: '26', value: 'date26' },
        { text: '27', value: 'date27' },
        { text: '28', value: 'date28' },
        { text: '29', value: 'date29' },
        { text: '30', value: 'date30' },
        { text: '31', value: 'date31' },
        { text: 'Кол-во дней за месяц', value: 'amountMonths' },
        { text: 'Кол-во часов за месяц', value: 'amountMonths' },
        { text: 'Кол-во ночных', value: 'amountMonths' },
        { text: 'Кол-во праздничных', value: 'amountMonths' },
        { text: 'Кол-во выходных', value: 'amountMonths' }
      ],
      items: [
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 },
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 },
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 },
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 },
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 },
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 },
        { name: 'fsdfsd', number: 4343, post: 'Директор', date1: '8', date2: '6', date3: '8', date4: '4', date5: '5', date6: 6, date7: 9, date8: 8, date9: 5, date10: 6, date11: 9, date12: 8, date13: 5, date14: 5, date15: 7 }
      ]
    }
  },
  methods: {
    save () {},
    close () {},
    open () {},
    cancel () {}
  }
}

</script>

<style scoped>

</style>
